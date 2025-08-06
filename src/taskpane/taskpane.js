Office.onReady(async () => {
  try {
    const token = await OfficeRuntime.storage.getItem("authToken");

    if (!token) {
      // No token found, redirect to login
      window.location.href = "login.html";
    } else {
      // Token found, show sidebar UI
      window.location.href = "sidebar.html";
    }
  } catch (err) {
    console.error("Error checking token:", err);
  }
});

document.getElementById("logout").addEventListener("click", async () => {
  try {
    await OfficeRuntime.storage.removeItem("authToken"); // ⛔ Remove token
    window.location.href = "login.html"; // 🔁 Redirect to login
  } catch (error) {
    console.error("Logout failed:", error);
  }
});

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
