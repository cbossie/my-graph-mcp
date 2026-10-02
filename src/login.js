#!/usr/bin/env node
// One-time interactive sign-in: npm run login
import { getToken, SCOPES } from "./auth.js";

await getToken({ interactive: true });
console.log(`Signed in. Scopes requested: ${SCOPES.join(" ")}`);
