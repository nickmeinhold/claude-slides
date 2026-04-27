import fs from "fs/promises";
import path from "path";
import os from "os";

export interface StoredTokens {
  access_token: string;
  refresh_token: string;
  expiry_date: number;
  scope: string;
}

const TOKEN_DIR = path.join(os.homedir(), ".claude-slides");
const TOKEN_FILE = path.join(TOKEN_DIR, "tokens.json");

export async function saveTokens(tokens: StoredTokens): Promise<void> {
  await fs.mkdir(TOKEN_DIR, { recursive: true });
  await fs.writeFile(TOKEN_FILE, JSON.stringify(tokens, null, 2), "utf-8");
  await fs.chmod(TOKEN_FILE, 0o600);
}

export async function loadTokens(): Promise<StoredTokens | null> {
  try {
    const data = await fs.readFile(TOKEN_FILE, "utf-8");
    return JSON.parse(data);
  } catch {
    return null;
  }
}

export async function clearTokens(): Promise<void> {
  try {
    await fs.unlink(TOKEN_FILE);
  } catch {
    // Ignore if file doesn't exist
  }
}
