/* Task pane UI wiring. */
import { smartFormatSelection } from "../core/smartFormat";
import { buildWaterfallFromSelection } from "../core/waterfall";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in runs in Excel.", true);
    return;
  }
  byId("btn-format").addEventListener("click", () => runAction(smartFormatSelection));
  byId("btn-waterfall").addEventListener("click", () => runAction(buildWaterfallFromSelection));
});

function byId(id: string): HTMLElement {
  const el = document.getElementById(id);
  if (!el) throw new Error(`missing #${id}`);
  return el;
}

function setStatus(msg: string, isError = false) {
  const s = byId("status");
  s.textContent = msg;
  s.className = "status " + (isError ? "err" : "ok");
}

async function runAction(fn: () => Promise<string>) {
  const buttons = document.querySelectorAll<HTMLButtonElement>("button.action");
  buttons.forEach((b) => (b.disabled = true));
  setStatus("Working…");
  try {
    const result = await fn();
    setStatus(result);
  } catch (e) {
    setStatus(e instanceof Error ? e.message : String(e), true);
  } finally {
    buttons.forEach((b) => (b.disabled = false));
  }
}
