// A tiny stand-in for Cloudflare KV, backed by Upstash Redis (add it in Vercel → Storage).
// The sign-in library expects KV, so this gives it the same get/put/delete/list shape.
import { Redis } from "@upstash/redis";

type PutOptions = { expirationTtl?: number; expiration?: number; metadata?: unknown };
type ListOptions = { prefix?: string; cursor?: string; limit?: number };
type ListKey = { name: string; expiration?: number; metadata?: unknown };
type ListResult = { keys: ListKey[]; list_complete: boolean; cursor?: string };

export interface KVLike {
  get(key: string, options?: { type?: "text" | "json" } | "text" | "json"): Promise<any>;
  put(key: string, value: string, options?: PutOptions): Promise<void>;
  delete(key: string): Promise<void>;
  list(options?: ListOptions): Promise<ListResult>;
}

const META = "__meta:"; // metadata lives in a sibling key so list() can return it

function ttlSeconds(o?: PutOptions): number | undefined {
  if (o?.expirationTtl) return Math.max(60, Math.ceil(o.expirationTtl));
  if (o?.expiration) return Math.max(60, Math.ceil(o.expiration - Date.now() / 1000));
  return undefined;
}

class RedisKV implements KVLike {
  constructor(private r: Redis) {}

  async get(key: string, options?: { type?: "text" | "json" } | "text" | "json") {
    const type = typeof options === "string" ? options : options?.type;
    const v = await this.r.get<string>(key);
    if (v === null || v === undefined) return null;
    return type === "json" ? JSON.parse(v) : v;
  }

  async put(key: string, value: string, options?: PutOptions) {
    const ex = ttlSeconds(options);
    const p = this.r.pipeline();
    ex ? p.set(key, value, { ex }) : p.set(key, value);
    if (options?.metadata !== undefined) {
      const m = JSON.stringify(options.metadata);
      ex ? p.set(META + key, m, { ex }) : p.set(META + key, m);
    } else {
      p.del(META + key);
    }
    await p.exec();
  }

  async delete(key: string) {
    await this.r.del(key, META + key);
  }

  async list(options: ListOptions = {}): Promise<ListResult> {
    const limit = Math.min(options.limit ?? 1000, 1000);
    const pattern = (options.prefix ?? "").replace(/[*?[\]\\]/g, "\\$&") + "*";
    let cursor = options.cursor ?? "0";
    const names: string[] = [];
    // SCAN may return fewer than asked; keep going until we have enough or reach the end.
    do {
      const [next, batch] = await this.r.scan(cursor, { match: pattern, count: limit });
      cursor = String(next);
      for (const k of batch as string[]) if (!k.startsWith(META)) names.push(k);
    } while (cursor !== "0" && names.length < limit);

    const metas = names.length ? await this.r.mget<(string | null)[]>(...names.map((n) => META + n)) : [];
    const keys = names.map((name, i) => ({
      name,
      ...(metas[i] ? { metadata: JSON.parse(metas[i] as string) } : {}),
    }));
    return cursor === "0" ? { keys, list_complete: true } : { keys, list_complete: false, cursor };
  }
}

// For local testing only: keeps everything in memory.
class MemoryKV implements KVLike {
  private m = new Map<string, { v: string; exp?: number; meta?: unknown }>();
  private live(k: string) {
    const e = this.m.get(k);
    if (e?.exp && e.exp < Date.now()) {
      this.m.delete(k);
      return undefined;
    }
    return e;
  }
  async get(key: string, options?: { type?: "text" | "json" } | "text" | "json") {
    const type = typeof options === "string" ? options : options?.type;
    const e = this.live(key);
    if (!e) return null;
    return type === "json" ? JSON.parse(e.v) : e.v;
  }
  async put(key: string, value: string, options?: PutOptions) {
    const ex = ttlSeconds(options);
    this.m.set(key, { v: value, exp: ex ? Date.now() + ex * 1000 : undefined, meta: options?.metadata });
  }
  async delete(key: string) {
    this.m.delete(key);
  }
  async list(options: ListOptions = {}) {
    const keys = [...this.m.keys()]
      .filter((k) => k.startsWith(options.prefix ?? "") && this.live(k))
      .sort()
      .map((name) => ({ name, ...(this.m.get(name)!.meta !== undefined ? { metadata: this.m.get(name)!.meta } : {}) }));
    return { keys, list_complete: true };
  }
}

let kv: KVLike | undefined;

export function getKV(): KVLike {
  if (kv) return kv;
  const url = process.env.KV_REST_API_URL ?? process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN ?? process.env.UPSTASH_REDIS_REST_TOKEN;
  if (url && token) {
    kv = new RedisKV(new Redis({ url, token, automaticDeserialization: false }));
  } else if (process.env.NODE_ENV !== "production" || process.env.ALLOW_MEMORY_KV === "1") {
    console.warn("No Upstash Redis found: using in-memory storage (testing only).");
    kv = new MemoryKV();
  } else {
    throw new Error("Storage is missing. In Vercel, add Upstash Redis under Storage and connect it to this project.");
  }
  return kv;
}
