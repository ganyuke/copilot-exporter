/*type CopilotMessage = {
  text: string;
  author: string; // either "bot" or "user"
  createdAt: string; // ISO-8601 timestamp
  timestamp: string; // ISO-8601 timestamp
  messageId: string; // uuid
  requestId: string; // uuid
  offense: string; // "None"
  sourceAttributions: {
    sourceType: string; // "CITATION"
    providerDisplayName: string;
    referenceMetadata: string; // JSON string
    isCitedInResponse: string; // boolean string
    imageLink: string; // url
    imageWidth: string; // number string
    imageHeight: string; // number string
    imageFavicon: string; // base64
    searchQuery: string;
    videoTimeLength: string;
    videoDAPublicationDate: string;
    videoViewCount: string;
    seeMoreUrl: string // url
  }[],
  contentOrigin: string; // "DeepLeo", "officeweb"
  turnCount: number; // index of message in conversation
  storageMessageId: string; // 13-number string  0: "text"
}

type CopilotBotMessage = CopilotMessage & {
  adaptiveCards: {
    type: string;
    version: string;
    body: {
      type: string;
      text: string;
      wrap: boolean;
    }[];
  }[];
  scores: {
    component: string; // "BotOffense"
    score: number; // decimal of 0.(9-11 numbers)
  }[];
  spokenText: string; // looks empty?
}

type CopilotUserMessage = CopilotMessage & {
  from: {
    id: string; // uuid
  };
  local: string; // locale string "en-US"
  market: string; // locale stirng "en-us"
  region: string; // locale string "us"
  locationInfo: {
    country: string; // "United States"
    state: string;
    city: string;
    timeZone: string; // Atlantic/Reykjavik
    timeZoneOffset: number;
    sourceType: number;
  };
  inputMethod: string; // "Keyboard"
  entityAnnotationTypes: string[]; // "People", "File", etc.
}*/

// bot messages in `/GetChats` have scores
type CopilotOverviewMessage = {
  text: string;
  author: string; // either "bot" or "user"
  responseIdentifier: string;
  createdAt: string; // ISO-8601 timestamp
  timestamp: string; // ISO-8601 timestamp
  messageId: string; // uuid
  requestId: string; // uuid
  offense: string; // "None"
  adaptiveCards: {
    type: string;
    version: string;
    body: {
      type: string;
      text: string;
      wrap: boolean;
    }[];
  }[];
  contentOrigin: string; // "DeepLeo"
  scores: {
    component: string; // "BotOffense"
    score: number; // decimal of 0.(9-11 numbers)
  }[];
  spokenText: string; // looks empty?
  turnCount: number; // index of message in conversation
  storageMessageId: string; // 13-number string
}

type CopilotConversationOverview = {
  conversationId: string; // uuid
  chatName: string;
  tone: string;
  createTimeUtc: number; // 13-number timestamp 
  updateTimeUtc: number; // 13-number timestamp 
  expiryTimeUtc: number; // workspace dependant?
  plugins: {
    id: string;
    source: string;
    isThirdPartyPluginSource: boolean;
    isGraphConnectorPluginType: boolean;
  }[];
  threadLevelGptId: {}; // not sure
  isMessageless: boolean;
  isUnread: boolean;
  retentionPolicyEffect: number;
  threadId: string;
  isScheduledPromptThread: boolean;
  lastMessage: CopilotOverviewMessage;
  mostRecentGptIds: []; // not sure
  hasLoopPages: boolean;
  isLegacyWebChat: boolean;
}

type CopilotChats = {
  chats: CopilotConversationOverview[];
  totalCountOfSavedChats: number;
  syncState: string;
  retentionPolicyStatus: number;
  result: {
    value: string;
    message: string;
    serviceVersion: string;
  }
}

export async function fetchCopilotChats(
  token: string,
  userOid: string,
  tenantId: string,
  maxChats: number,
  variants: string = 'feature.EnableLastMessageForGetChats,feature.EnableMRUAgents,feature.EnableHasLoopPages'
): Promise<CopilotChats> {
  const requestObj = {
    source: "officeweb",
    traceId: crypto.randomUUID(), // uuid with spaces
    threadType: "webchat",
    MaxReturnedChatsCount: maxChats
  };

  const encodedRequest = encodeURIComponent(JSON.stringify(requestObj));
  const encodedVariants = encodeURIComponent(variants);

  const url = `https://substrate.office.com/m365Copilot/GetChats?request=${encodedRequest}&variants=${encodedVariants}`;

  const headers = {
    "authorization": `Bearer ${token}`,
    "content-type": "application/json",
    "x-anchormailbox": `Oid:${userOid}@${tenantId}`,
    "x-clientrequestid": crypto.randomUUID().replace(/-/g, ""), // uuid *without* spaces
    "x-routingparameter-sessionkey": userOid,
    "x-scenario": "OfficeWebIncludedCopilot",
    "x-variants": variants
  };

  const res = await fetch(url, {
    method: "GET",
    headers
  });

  if (!res.ok) {
    throw new Error(`Fetch failed with status ${res.status}`);
  }

  const data = await res.json() as CopilotChats;
  return data;
}

export async function fetchCopilotConversation(
  token: string,
  userOid: string,
  tenantId: string,
  conversationId: string
) {
  const requestObj = {
    conversationId,
    source: "officeweb",
    traceId: crypto.randomUUID().replace(/-/g, ""), // uuid *without* spaces (for some reason??)
  };

  const encodedRequest = encodeURIComponent(JSON.stringify(requestObj));

  const url = `https://substrate.office.com/m365Copilot/GetConversation?request=${encodedRequest}`;

  const headers = {
    "authorization": `Bearer ${token}`,
    "content-type": "application/json",
    "x-anchormailbox": `Oid:${userOid}@${tenantId}`,
    "x-clientrequestid": crypto.randomUUID().replace(/-/g, ""), // also UUID w/o spaces
    "x-routingparameter-sessionkey": userOid,
    "x-scenario": "OfficeWebIncludedCopilot"
  };

  const response = await fetch(url, {
    method: "GET",
    headers
  });

  if (!response.ok) {
    throw new Error(`Fetch failed with status ${response.status}`);
  }

  return await response.blob();
}
