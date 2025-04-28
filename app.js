// MSAL configuration for user authentication
const msalConfig = {
  auth: {
    clientId: "167a1b7b-50ed-4a39-bd23-f93110dcdcf3", // ← Serve SEMPRE!
    authority: "https://login.microsoftonline.com/common",  // Multi-tenant authority
    redirectUri: window.location.href, // URL to redirect after login
  },
  cache: {
    cacheLocation: "sessionStorage", // Store session data in the browser
    storeAuthStateInCookie: true, // Store state in cookies if needed
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

let token = null;

// Function to get the access token from the logged-in user
async function getAccessToken() {
  try {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("No account found");
    }

    const request = {
      account: accounts[0],
      scopes: ["https://graph.microsoft.com/.default"], // Scope for SharePoint access
    };

    const response = await msalInstance.acquireTokenSilent(request);
    return response.accessToken;
  } catch (error) {
    console.error("Error fetching token:", error);
    return null;
  }
}

// Function to get SharePoint sites the user has access to
async function getSharePointSites(token) {
  const client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token); // Provide the access token
    },
  });

  try {
    const response = await client.api("/me/sites").get(); // Get SharePoint sites
    return response.value; // Return the list of SharePoint sites
  } catch (error) {
    console.error("Error fetching SharePoint sites:", error);
    return [];
  }
}

// Function to get lists from a specific SharePoint site
async function getSharePointLists(token, siteId) {
  const client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token); // Provide the access token
    },
  });

  try {
    const response = await client.api(`/sites/${siteId}/lists`).get(); // Get lists from a specific site
    return response.value; // Return the list of lists
  } catch (error) {
    console.error("Error fetching SharePoint lists:", error);
    return [];
  }
}

// Function to get items from a SharePoint list
async function getListItems(token, siteId, listId) {
  const client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, token); // Provide the access token
    },
  });

  try {
    const response = await client.api(`/sites/${siteId}/lists/${listId}/items`).get(); // Get list items
    return response.value; // Return the list items
  } catch (error) {
    console.error("Error fetching list items:", error);
    return [];
  }
}

// Function to handle user queries and fetch answers from SharePoint
async function handleUserQuery(query) {
  const answerElement = document.getElementById('answer');

  // Get the access token
  token = await getAccessToken();
  if (!token) {
    answerElement.innerHTML = "Error retrieving token.";
    return;
  }

  // Fetch SharePoint sites
  const sites = await getSharePointSites(token);
  if (sites.length === 0) {
    answerElement.innerHTML = "No SharePoint sites found.";
    return;
  }

  // Fetch lists from the first site
  const siteId = sites[0].id;
  const lists = await getSharePointLists(token, siteId);
  if (lists.length === 0) {
    answerElement.innerHTML = "No SharePoint lists found.";
    return;
  }

  // Fetch list items from the first list
  const listId = lists[0].id;
  const items = await getListItems(token, siteId, listId);
  if (items.length === 0) {
    answerElement.innerHTML = "No items found in the list.";
    return;
  }

  // Search for an answer from the list items (for example, by title)
  const answer = items.find(item => item.fields.Title.includes(query));
  if (answer) {
    answerElement.innerHTML = `Answer found: ${answer.fields.Description}`;
  } else {
    answerElement.innerHTML = "No answer found.";
  }
}

// Event listener for login button to trigger MSAL login
document.getElementById('login-button').addEventListener('click', () => {
  msalInstance.loginPopup().then(() => {
    document.getElementById('login-section').style.display = 'none';
    document.getElementById('question-section').style.display = 'block';
  });
});

// Event listener for asking a question
document.getElementById('ask-button').addEventListener('click', () => {
  const query = document.getElementById('question').value;
  if (query) {
    handleUserQuery(query);
  } else {
    document.getElementById('answer').innerHTML = "Please enter a question.";
  }
});
