// ===== 設定 =====
const GAS_URL = "https://script.google.com/macros/s/AKfycbzDv8ZFVd3NFv8YKhNRHcatoVQ9u_T4AHYtGrF0ALrUOVzicAvZ2RSv5YCu-uVNocmZJw/exec";
const DEMO_MODE = false; // 展示模式：true=使用假資料 / false=連接GAS

// ===== 安全性與工具函式 =====
function escapeHtml(text) {
  if (!text) return "";
  const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

// 格式化貨幣
function formatCurrency(num) {
  return Number(num || 0).toLocaleString();
}

// ===== 登入驗證 =====
function checkAuth() {
  if (sessionStorage.getItem("auth") !== "1") {
    location.href = "index.html";
  }
}

// ===== Modal 關閉 =====
function closeModal(id) {
  document.getElementById(id).style.display = "none";
}

// 點擊遮罩關閉 Modal
document.addEventListener("click", e => {
  if (e.target.classList.contains("modal-overlay")) {
    e.target.style.display = "none";
  }
});

// ===== UI 提示控制 =====
let activeRequests = 0;
function showLoading() {
  activeRequests++;
  let loader = document.getElementById("global-loader");
  if (!loader) {
    loader = document.createElement("div");
    loader.id = "global-loader";
    loader.innerHTML = '<div class="spinner"></div>';
    document.body.appendChild(loader);
  }
  loader.style.display = "flex";
}

function hideLoading() {
  activeRequests--;
  if (activeRequests > 0) return;
  activeRequests = 0;
  const loader = document.getElementById("global-loader");
  if (loader) loader.style.display = "none";
}

function showToast(msg, isError = false) {
  const toast = document.createElement("div");
  toast.className = `toast ${isError ? "error" : "success"}`;
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 300);
    }, 2500);
  }, 100);
}

// ===== GAS 通訊 =====
function fetchData(action, params, callback) {
  if (DEMO_MODE) {
    console.log("DEMO_MODE: ", action, params);
    callback({ success: false, data: [] });
    return;
  }
  showLoading();
  const url = `${GAS_URL}?action=${action}&${new URLSearchParams(params)}`;
  fetch(url)
    .then(r => r.json())
    .then(res => {
      hideLoading();
      if (res.error) showToast(res.error, true);
      callback(res);
    })
    .catch(err => {
      hideLoading();
      showToast("連線失敗，請檢查網路", true);
      callback({ success: false, data: [] });
    });
}
