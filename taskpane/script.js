Office.onReady(() => {
  // Ensure Office is ready before using APIs
  document.getElementById("fetchEmails").onclick = getUnreadEmails;
  document.getElementById("summarizeSelected").onclick = summarizeSelected;
});

function getUnreadEmails() {
  Office.auth.getAccessTokenAsync({ allowSignInPrompt: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const token = result.value;

      fetch("https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false&$top=10", {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        }
      })
      .then(res => res.json())
      .then(data => {
        const emailList = document.getElementById("emailList");
        emailList.innerHTML = "";

        data.value.forEach((email, index) => {
          const body = email.body.content.replace(/(<([^>]+)>)/gi, ""); // Strip HTML
          emailList.innerHTML += `
            <div>
              <input type="checkbox" id="email${index}" value="${encodeURIComponent(body)}">
              <label for="email${index}"><strong>${email.subject}</strong><br>${body.substring(0, 100)}...</label>
            </div><br>
          `;
        });
      })
      .catch(err => {
        console.error("Error fetching emails:", err);
        alert("Failed to fetch emails.");
      });
    } else {
      console.error("SSO token error:", result.error.message);
      alert("Authentication failed. Please sign in again.");
    }
  });
}

function summarizeSelected() {
  const selectedEmails = [...document.querySelectorAll("input[type=checkbox]:checked")].map(e =>
    decodeURIComponent(e.value)
  );

  if (selectedEmails.length === 0) {
    alert("Please select at least one email.");
    return;
  }

  fetch("https://your-backend-url.com/summarize", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ emails: selectedEmails })
  })
    .then(res => res.json())
    .then(data => {
      let output = "";
      data.forEach((item, i) => {
        output += `📧 Email ${i + 1}:\n`;
        output += `📝 Summary: ${item.summary}\n`;
        output += `⚠️ Urgency: ${item.urgency}\n`;
        output += `✅ Action: ${item.action}\n\n`;
      });
      alert(output);
    })
    .catch(err => {
      console.error("Error summarizing emails:", err);
      alert("Failed to summarize emails.");
    });
}
