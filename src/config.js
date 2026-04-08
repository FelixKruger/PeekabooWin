import os from "node:os";
import path from "node:path";

export function getAppHome() {
  return (
    process.env.PEEKABOO_WINDOWS_HOME ??
    path.join(os.homedir(), ".peekaboo-windows")
  );
}

export function getSnapshotsDir() {
  return path.join(getAppHome(), "snapshots");
}

export function getHarvestsDir() {
  return path.join(getAppHome(), "harvests");
}
