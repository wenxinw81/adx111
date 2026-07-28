const form = document.querySelector("#reportForm");
const input = document.querySelector("#promptInput");
const parsedDate = document.querySelector("#parsedDate");
const parsedType = document.querySelector("#parsedType");
const parsedOrder = document.querySelector("#parsedOrder");
const jobStatus = document.querySelector("#jobStatus");
const messageBox = document.querySelector("#messageBox");
const downloadLink = document.querySelector("#downloadLink");
const progressBar = document.querySelector("#progressBar");
const envStatus = document.querySelector("#envStatus");
const submitButton = form.querySelector("button[type='submit']");

const statusText = {
  queued: "排队中",
  running: "生成中",
  success: "已完成",
  failed: "失败",
};

function setProgress(status) {
  const width = status === "success" ? 100 : status === "running" ? 72 : status === "queued" ? 28 : 0;
  progressBar.style.width = `${width}%`;
}

function renderJob(job) {
  parsedDate.textContent = job.request.report_date;
  parsedType.textContent = job.request.analysis_name;
  parsedOrder.textContent = job.request.order_id ? `订单${job.request.order_id}` : "全部";
  jobStatus.textContent = statusText[job.status] || job.status;
  envStatus.textContent = statusText[job.status] || job.status;
  setProgress(job.status);

  if (job.status === "success") {
    messageBox.className = "message";
    messageBox.textContent = `报告已生成：\n${job.output_path}`;
    downloadLink.href = job.download_url;
    downloadLink.classList.remove("hidden");
    submitButton.disabled = false;
    return true;
  }

  if (job.status === "failed") {
    messageBox.className = "message failed";
    messageBox.textContent = job.error || "生成失败，请检查数据库连接和环境变量。";
    downloadLink.classList.add("hidden");
    submitButton.disabled = false;
    return true;
  }

  messageBox.className = "message warning";
  messageBox.textContent = job.status === "queued" ? "任务已提交，等待执行。" : "正在连接数据库并生成 Excel，请稍等。";
  downloadLink.classList.add("hidden");
  return false;
}

async function pollJob(jobId) {
  for (;;) {
    const response = await fetch(`/api/reports/${jobId}`);
    const job = await response.json();
    const done = renderJob(job);
    if (done) break;
    await new Promise((resolve) => setTimeout(resolve, 1800));
  }
}

async function submitReport(text) {
  submitButton.disabled = true;
  downloadLink.classList.add("hidden");
  progressBar.style.width = "10%";
  messageBox.className = "message warning";
  messageBox.textContent = "正在解析需求。";

  const response = await fetch("/api/reports", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });

  if (!response.ok) {
    submitButton.disabled = false;
    messageBox.className = "message failed";
    messageBox.textContent = await response.text();
    return;
  }

  const job = await response.json();
  renderJob(job);
  await pollJob(job.job_id);
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  const text = input.value.trim();
  if (!text) return;
  submitReport(text).catch((error) => {
    submitButton.disabled = false;
    messageBox.className = "message failed";
    messageBox.textContent = error.message;
  });
});

document.querySelectorAll("[data-text]").forEach((button) => {
  button.addEventListener("click", () => {
    input.value = button.dataset.text;
    input.focus();
  });
});
