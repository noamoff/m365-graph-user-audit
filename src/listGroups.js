/**
 * Microsoft 365 Graph API – List Groups
 */

import { Client } from "@microsoft/microsoft-graph-client";

export async function listGroups(accessToken) {
  const client = Client.init({
    authProvider: done => done(null, accessToken)
  });

  const response = await client
    .api("/groups")
    .select("id,displayName")
    .get();

  return response.value;
}
