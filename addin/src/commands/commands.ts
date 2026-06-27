/* Ribbon button handlers (ExecuteFunction actions). */
import { smartFormatSelection } from "../core/smartFormat";
import { buildWaterfallFromSelection } from "../core/waterfall";

Office.onReady(() => {});

async function run(fn: () => Promise<string>, event: Office.AddinCommands.Event) {
  try {
    await fn();
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    console.error(msg);
  } finally {
    event.completed();
  }
}

function smartFormatTable(event: Office.AddinCommands.Event) {
  run(smartFormatSelection, event);
}

function buildWaterfall(event: Office.AddinCommands.Event) {
  run(buildWaterfallFromSelection, event);
}

// Register with Office so the manifest FunctionName resolves.
Office.actions.associate("smartFormatTable", smartFormatTable);
Office.actions.associate("buildWaterfall", buildWaterfall);
