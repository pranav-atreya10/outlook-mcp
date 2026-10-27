// Sends every request on this site to the connector.
import { handle } from "../../lib/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

export const GET = (req: Request) => handle(req);
export const POST = (req: Request) => handle(req);
export const DELETE = (req: Request) => handle(req);
export const OPTIONS = (req: Request) => handle(req);
