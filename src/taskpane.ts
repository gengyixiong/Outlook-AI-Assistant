type CheckState = "pass" | "fail" | "unknown";

interface CapabilityResult {
  name: string;
  state: CheckState;
  value: string;
}

interface MailboxDiagnostics {
  host: string;
  platform: string;
  version: string;
  mailbox15: boolean;
  mailbox110: boolean;
  mailbox114: boolean;
  nestedAppAuth11: boolean;
  itemChangedRegistered: boolean;
}

const capabilityList = document.querySelector<HTMLDivElement>("#capability-list");
const lastChecked = document.querySelector<HTMLParagraphElement>("#last-checked");
const messageFields = document.querySelector<HTMLDListElement>("#message-fields");
const bodyPreviewPanel = document.querySelector<HTMLDetailsElement>("#body-preview-panel");
const bodyPreview = document.querySelector<HTMLPreElement>("#body-preview");
const eventLog = document.querySelector<HTMLDivElement>("#event-log");
const runChecksButton = document.querySelector<HTMLButtonElement>("#run-checks");
const readCurrentButton = document.querySelector<HTMLButtonElement>("#read-current");

let itemChangedRegistered = false;

function addLog(message: string): void {
  if (!eventLog) {
    return;
  }

  const item = document.createElement("div");
  item.className = "event-log__item";

  const time = document.createElement("span");
  time.className = "event-log__time";
  time.textContent = new Date().toLocaleTimeString("zh-CN", { hour12: false });

  const text = document.createElement("span");
  text.textContent = message;

  item.append(time, text);
  eventLog.prepend(item);
}

function isRequirementSupported(name: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(name, version);
  } catch {
    return false;
  }
}

function getDiagnostics(): MailboxDiagnostics {
  const diagnostics = Office.context.diagnostics;

  return {
    host: String(diagnostics?.host ?? "Outlook"),
    platform: String(diagnostics?.platform ?? "未知"),
    version: String(diagnostics?.version ?? "未知"),
    mailbox15: isRequirementSupported("Mailbox", "1.5"),
    mailbox110: isRequirementSupported("Mailbox", "1.10"),
    mailbox114: isRequirementSupported("Mailbox", "1.14"),
    nestedAppAuth11: isRequirementSupported("NestedAppAuth", "1.1"),
    itemChangedRegistered
  };
}

function boolCheck(name: string, value: boolean, detail: string): CapabilityResult {
  return {
    name,
    state: value ? "pass" : "fail",
    value: value ? detail : "不支持"
  };
}

function renderCapabilities(): void {
  if (!capabilityList || !lastChecked) {
    return;
  }

  const diagnostics = getDiagnostics();
  const checks: CapabilityResult[] = [
    {
      name: "Outlook 主机",
      state: "pass",
      value: `${diagnostics.platform} · ${diagnostics.version}`
    },
    boolCheck("Mailbox 1.5", diagnostics.mailbox15, "支持固定窗格基础能力"),
    boolCheck("Mailbox 1.10", diagnostics.mailbox110, "支持较新事件能力"),
    boolCheck("Mailbox 1.14", diagnostics.mailbox114, "符合 LTSC 2024 预期"),
    boolCheck("NestedAppAuth 1.1", diagnostics.nestedAppAuth11, "可用于 Graph 登录"),
    boolCheck("ItemChanged 监听", diagnostics.itemChangedRegistered, "已注册")
  ];

  capabilityList.replaceChildren(
    ...checks.map((check) => {
      const row = document.createElement("div");
      row.className = `capability capability--${check.state}`;

      const icon = document.createElement("span");
      icon.className = "capability__icon";
      icon.textContent = check.state === "pass" ? "✓" : check.state === "fail" ? "!" : "?";

      const name = document.createElement("span");
      name.className = "capability__name";
      name.textContent = check.name;

      const value = document.createElement("span");
      value.className = "capability__value";
      value.textContent = check.value;

      row.append(icon, name, value);
      return row;
    })
  );

  lastChecked.textContent = `最近检测：${new Date().toLocaleString("zh-CN")}`;
  addLog("兼容性检测完成。");
}

function formatAddress(
  recipient: Office.EmailAddressDetails | undefined
): string {
  if (!recipient) {
    return "未提供";
  }

  if (recipient.displayName && recipient.emailAddress) {
    return `${recipient.displayName} <${recipient.emailAddress}>`;
  }

  return recipient.displayName || recipient.emailAddress || "未提供";
}

function formatRecipientList(
  recipients: Office.EmailAddressDetails[] | undefined
): string {
  if (!recipients?.length) {
    return "无";
  }

  return recipients.map(formatAddress).join("; ");
}

function setMessageRows(rows: Array<[string, string]>): void {
  if (!messageFields) {
    return;
  }

  messageFields.replaceChildren(
    ...rows.map(([label, value]) => {
      const row = document.createElement("div");
      const term = document.createElement("dt");
      const detail = document.createElement("dd");
      term.textContent = label;
      detail.textContent = value;
      row.append(term, detail);
      return row;
    })
  );
}

function readCurrentMessage(): void {
  const item = Office.context.mailbox.item;

  if (!item) {
    setMessageRows([["状态", "当前没有选中邮件。"]]);
    if (bodyPreviewPanel) {
      bodyPreviewPanel.hidden = true;
    }
    addLog("没有可读取的当前邮件。");
    return;
  }

  const message = item as Office.MessageRead;
  const subject = message.subject || "（无主题）";
  const from = formatAddress(message.from);
  const to = formatRecipientList(message.to);
  const cc = formatRecipientList(message.cc);
  const receivedTime =
    message.dateTimeCreated instanceof Date
      ? message.dateTimeCreated.toLocaleString("zh-CN")
      : "未提供";
  const attachmentCount = String(message.attachments?.length ?? 0);

  setMessageRows([
    ["主题", subject],
    ["发件人", from],
    ["收件人", to],
    ["抄送", cc],
    ["时间", receivedTime],
    ["附件数量", attachmentCount],
    ["正文", "正在读取…"]
  ]);

  message.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setMessageRows([
        ["主题", subject],
        ["发件人", from],
        ["状态", `正文读取失败：${result.error?.message ?? "未知错误"}`]
      ]);
      if (bodyPreviewPanel) {
        bodyPreviewPanel.hidden = true;
      }
      addLog("当前邮件正文读取失败。");
      return;
    }

    const text = result.value ?? "";
    setMessageRows([
      ["主题", subject],
      ["发件人", from],
      ["收件人", to],
      ["抄送", cc],
      ["时间", receivedTime],
      ["附件数量", attachmentCount],
      ["正文长度", `${text.length.toLocaleString("zh-CN")} 个字符`]
    ]);

    if (bodyPreview && bodyPreviewPanel) {
      bodyPreview.textContent = text.slice(0, 500).trim() || "（正文为空）";
      bodyPreviewPanel.hidden = false;
    }

    addLog(`已读取当前邮件：${subject.slice(0, 40)}`);
  });
}

function handleItemChanged(): void {
  addLog("检测到邮件切换事件。");
  renderCapabilities();
  readCurrentMessage();
}

function registerItemChanged(): Promise<void> {
  return new Promise((resolve) => {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      handleItemChanged,
      (result) => {
        itemChangedRegistered =
          result.status === Office.AsyncResultStatus.Succeeded;

        if (itemChangedRegistered) {
          addLog("ItemChanged 事件监听注册成功。");
        } else {
          addLog(`ItemChanged 注册失败：${result.error?.message ?? "未知错误"}`);
        }

        resolve();
      }
    );
  });
}

runChecksButton?.addEventListener("click", renderCapabilities);
readCurrentButton?.addEventListener("click", readCurrentMessage);

Office.onReady(async (info) => {
  addLog(`Office.js 已初始化：${info.host ?? "未知主机"}`);
  await registerItemChanged();
  renderCapabilities();
  readCurrentMessage();
});
