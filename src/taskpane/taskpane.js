Office.onReady(async () => {
  try {
    const token = await OfficeRuntime.storage.getItem("authToken");
    const btn = document.getElementById("xero-login-btn");
    if (btn) {
      btn.addEventListener("click", loginWithXero);
    }
    checkTokenAndRedirect(token); // Check token and handle redirection
  } catch (err) {
    console.error("Error checking token:", err);
  }
});

document.getElementById("logout").addEventListener("click", async () => {
  try {
    await OfficeRuntime.storage.removeItem("authToken"); // Remove token
    checkTokenAndRedirect(null); // Re-check the token and redirect to login
  } catch (error) {
    console.error("Logout failed:", error);
  }
});

function checkTokenAndRedirect(token) {
  if (!token) {
    // No token found, redirect to login if not already there
    if (window.location.pathname !== "/login.html") {
      window.location.href = "login.html";
    }
  } else {
    // Token found, redirect to sidebar if not already there
    if (window.location.pathname !== "/sidebar.html") {
      window.location.href = "sidebar.html";
    }
  }
}

function renderSidebarUI() {
  const app = document.getElementById("app");
  app.innerHTML = `
    <h2>Welcome to Accounting BI</h2>
    <p>You are logged in. 🎉</p>
    <button id="logout">Log out</button>
  `;

  document.getElementById("logout").onclick = async () => {
    await OfficeRuntime.storage.removeItem("authToken");
    location.reload();
  };
}

function loginWithXero() {
  console.log("calling this thing");

  Office.context.ui.displayDialogAsync(
    "https://staging.accounting.bi/login?redirect_url=true",
    { height: 60, width: 90, displayInIframe: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed to open:", asyncResult.error);
        return;
      }

      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        const token = arg.message; // Capture token from dialog
        console.log("Token received from dialog:", token);

        // Store the token in OfficeRuntime storage
        OfficeRuntime.storage.setItem("authToken", token).then(() => {
          // After storing the token, check for the token and redirect
          checkTokenAndRedirect(token); // Now token is stored, check again and redirect if needed
        });

        dialog.close(); // Close the dialog after storing the token
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        console.log("Dialog closed without login");
      });
    }
  );
}
