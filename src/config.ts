import { join } from "node:path";
const pkg = new URL("..", import.meta.url).pathname;
export const ROOT = process.env.PROPOSAL_ROOT ?? join(pkg, "samples");
export const WORKSPACE = process.env.PROPOSAL_WORKSPACE ?? join(pkg, "workspace");
export const STATE = join(WORKSPACE, "state.json");
