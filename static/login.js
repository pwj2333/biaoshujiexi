const $ = (id) => document.getElementById(id);

function showToast(message, isError = false) {
  const toast = $('toast');
  toast.textContent = message;
  toast.className = `toast ${isError ? 'error' : ''}`;
  setTimeout(() => {
    toast.className = 'toast hidden';
  }, 2600);
}

async function request(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    let message = '请求失败';
    try {
      const data = await response.json();
      message = data.detail || message;
    } catch (_) {}
    throw new Error(message);
  }
  return response.json();
}

async function login() {
  const result = await request('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      username: $('loginUsername').value.trim(),
      password: $('loginPassword').value,
    }),
  });
  localStorage.setItem('bid_parser_token', result.token);
  window.location.href = '/app';
}

function init() {
  $('loginBtn').onclick = async () => {
    try {
      await login();
    } catch (error) {
      showToast(error.message, true);
    }
  };
  $('loginPassword').addEventListener('keydown', async (event) => {
    if (event.key !== 'Enter') return;
    try {
      await login();
    } catch (error) {
      showToast(error.message, true);
    }
  });
}

init();
