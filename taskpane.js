// ----- taskpane.js -----
const taskpaneJs = `Office.onReady(() => {
  document.getElementById('generate').onclick = generateReply;
});

async function generateReply() {
  const apiKey = document.getElementById('apiKey').value.trim();
  if (!apiKey) { showStatus('Please enter your OpenAI API key.'); return; }
  showStatus('Fetching email content...');

  Office.context.mailbox.item.body.getAsync("text", async result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) { showStatus('Error retrieving email body.'); return; }
    const content = result.value;
    showStatus('Calling ChatGPT API...');

    try {
      const response = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: 'gpt-4',
          messages: [
            { role: 'system', content: 'You are an email assistant.' },
            { role: 'user', content: content }
          ],
          max_tokens: 500,
          temperature: 0.7
        })
      });
      const data = await response.json();
      const reply = data.choices[0].message.content;
      Office.context.mailbox.item.body.setSelectedDataAsync(reply, { coercionType: Office.CoercionType.Html }, res => {
        showStatus(res.status === Office.AsyncResultStatus.Succeeded ? 'Reply inserted successfully!' : 'Failed to insert reply.');
      });
    } catch (err) {
      showStatus('Error calling OpenAI: ' + err.message);
    }
  });
}

function showStatus(msg) {
  document.getElementById('status').innerText = msg;
}`;