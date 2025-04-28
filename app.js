// MSAL configuration
const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID", 
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Or "common"
    redirectUri: window.location.href
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let token = null;

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return null;

  try {
    const response = await msalInstance.acquireTokenSilent({
      account: accounts[0],
      scopes: ["https://graph.microsoft.com/.default"]
    });
    return response.accessToken;
  } catch {
    const response = await msalInstance.acquireTokenPopup({
      scopes: ["https://graph.microsoft.com/.default"]
    });
    return response.accessToken;
  }
}

function graphClient(token) {
  return MicrosoftGraph.Client.init({
    authProvider: (done) => done(null, token)
  });
}

async function getSharePointSites(token) {
  try {
    const client = graphClient(token);
    const response = await client.api("/sites?search=*").get();
    return response.value;
  } catch (e) {
    console.error("Sites error", e);
    return [];
  }
}

async function getSharePointLists(token, siteId) {
  try {
    const client = graphClient(token);
    const response = await client.api(`/sites/${siteId}/lists`).get();
    return response.value;
  } catch (e) {
    console.error("Lists error", e);
    return [];
  }
}

async function getListItems(token, siteId, listId) {
  try {
    const client = graphClient(token);
    const response = await client.api(`/sites/${siteId}/lists/${listId}/items?expand=fields`).get();
    return response.value;
  } catch (e) {
    console.error("Items error", e);
    return [];
  }
}

async function callLLM(question, items) {
  const context = items.map(i => `• ${i.fields?.Title || "No Title"}: ${i.fields?.Description || "No Description"}`).join('\n');

  const prompt = `
Answer this question using the following SharePoint data.

Question: "${question}"

Context:
${context}

Answer:
  `;

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer YOUR_OPENAI_API_KEY"
      },
      body: JSON.stringify({
        model: "gpt-3.5-turbo",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3,
        max_tokens: 300
      })
    });

    const data = await response.json();
    return data.choices?.[0]?.message?.content || "No answer generated.";
  } catch (error) {
    console.error("LLM error:", error);
    return "Error processing your question.";
  }
}

async function handleUserQuery(query) {
  const answerEl = document.getElementById("answer");
  answerEl.innerHTML = "Thinking...";

  token = await getAccessToken();
  if (!token) return (answerEl.innerHTML = "Authentication failed.");

  const sites = await getSharePointSites(token);
  if (!sites.length) return (answerEl.innerHTML = "No SharePoint sites found.");

  const siteId = sites[0].id;
  const lists = await getSharePointLists(token, siteId);
  if (!lists.length) return (answerEl.innerHTML = "No lists found.");

  const listId = lists[0].id;
  const items = await getListItems(token, siteId, listId);
  if (!items.length) return (answerEl.innerHTML = "No items found.");

  const answer = await callLLM(query, items);
  answerEl.innerHTML = answer;
}

// UI Events
document.getElementById("login-button").addEventListener("click", async () => {
  try {
    await msalInstance.loginPopup({ scopes: ["User.Read"] });
    document.getElementById("login-section").style.display = "none";
    document.getElementById("question-section").style.display = "block";
  } catch (e) {
    console.error("Login failed", e);
  }
});

document.getElementById("ask-button").addEventListener("click", () => {
  const query = document.getElementById("question").value;
  if (!query) {
    document.getElementById("answer").innerHTML = "Please enter a question.";
    return;
  }
  handleUserQuery(query);
});
