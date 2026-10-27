// Lets tests run outside Next.js by mapping "cloudflare:workers" to our stand-in.
import { register } from "node:module";

register(
  "data:text/javascript," +
    encodeURIComponent(`
      export async function resolve(specifier, context, next) {
        if (specifier === "cloudflare:workers") {
          return { url: "data:text/javascript,export class WorkerEntrypoint {}", shortCircuit: true };
        }
        return next(specifier, context);
      }`),
);
