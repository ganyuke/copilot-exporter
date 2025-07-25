export async function fetchCopilotChats(
  token: string,
  userOid: string,
  tenantId: string,
  variants: string = 'feature.EnableLastMessageForGetChats,feature.EnableMRUAgents,feature.EnableHasLoopPages'
) {
  const traceId = crypto.randomUUID().toUpperCase();          // UUID with spaces (we'll insert manually)
  const requestObj = {
    source: "officeweb",
    traceId: traceId.replace(/-/g, " "),                      // Add spaces to UUID for this param
    threadType: "webchat",
    MaxReturnedChatsCount: 40
  };

  const encodedRequest = encodeURIComponent(JSON.stringify(requestObj));
  const encodedVariants = encodeURIComponent(variants);

  const url = `https://substrate.office.com/m365Copilot/GetChats?request=${encodedRequest}&variants=${encodedVariants}`;

  const headers = {
    "authorization": `Bearer ${token}`,
    "content-type": "application/json",
    "x-anchormailbox": `Oid:${userOid}@${tenantId}`,
    "x-clientrequestid": crypto.randomUUID(),                // UUID without spaces
    "x-routingparameter-sessionkey": userOid,
    "x-scenario": "OfficeWebIncludedCopilot",
    "x-variants": variants
  };

  const res = await fetch(url, {
    method: "GET",
    headers,
    referrer: "https://m365.cloud.microsoft/"
  });

  if (!res.ok) {
    throw new Error(`Fetch failed with status ${res.status}`);
  }

  const data = await res.json();
  return data;
}

export async function fetchCopilotConversation(
  token: string,
  userOid: string,
  tenantId: string,
  conversationId: string
) {
  const traceId = crypto.randomUUID().replace(/-/g, ''); // UUID without spaces

  const requestObj = {
    conversationId,
    source: "officeweb",
    traceId
  };

  const encodedRequest = encodeURIComponent(JSON.stringify(requestObj));

  const url = `https://substrate.office.com/m365Copilot/GetChats?request=${encodedRequest}`;

  const headers = {
    "authorization": `Bearer ${token}`,
    "content-type": "application/json",
    "x-anchormailbox": `Oid:${userOid}@${tenantId}`,
    "x-clientrequestid": crypto.randomUUID().replace(/-/g, ''), // also UUID w/o spaces
    "x-routingparameter-sessionkey": userOid,
    "x-scenario": "OfficeWebIncludedCopilot"
  };

  const response = await fetch(url, {
    method: "GET",
    headers,
    referrer: "https://m365.cloud.microsoft/"
  });

  if (!response.ok) {
    throw new Error(`Fetch failed with status ${response.status}`);
  }

  return await response.json();
}
