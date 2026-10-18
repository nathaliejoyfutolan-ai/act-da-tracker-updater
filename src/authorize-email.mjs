import crypto from "node:crypto";
import fs from "node:fs";

const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "AUTH_PUBLIC_KEY_B64"];
for (const key of required) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const scope = "offline_access Files.ReadWrite User.Read Mail.Send";
const deviceResponse = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/devicecode`, {
  method: "POST",
  headers: { "Content-Type": "application/x-www-form-urlencoded" },
  body: new URLSearchParams({ client_id: process.env.MS_CLIENT_ID, scope })
});
if (!deviceResponse.ok) throw new Error(`Could not start Microsoft authorization (${deviceResponse.status}): ${await deviceResponse.text()}`);

const device = await deviceResponse.json();
console.log("\nMICROSOFT AUTHORIZATION REQUIRED");
console.log(device.message);
console.log("Keep this workflow running while authorization is completed.\n");

let intervalSeconds = Number(device.interval ?? 5);
const deadline = Date.now() + Number(device.expires_in ?? 900) * 1000;
let token;

while (Date.now() < deadline) {
  await delay(intervalSeconds * 1000);
  const response = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      client_id: process.env.MS_CLIENT_ID,
      device_code: device.device_code
    })
  });
  const result = await response.json();
  if (response.ok) {
    token = result;
    break;
  }
  if (result.error === "authorization_pending") continue;
  if (result.error === "slow_down") {
    intervalSeconds += 5;
    continue;
  }
  throw new Error(`Microsoft authorization failed: ${result.error_description ?? result.error}`);
}

if (!token?.refresh_token) throw new Error("Microsoft authorization expired before a refresh token was issued.");

const publicKey = crypto.createPublicKey(Buffer.from(process.env.AUTH_PUBLIC_KEY_B64, "base64").toString("utf8"));
const encryptionKey = crypto.randomBytes(32);
const iv = crypto.randomBytes(12);
const cipher = crypto.createCipheriv("aes-256-gcm", encryptionKey, iv);
const ciphertext = Buffer.concat([cipher.update(token.refresh_token, "utf8"), cipher.final()]);
const payload = {
  encryptedKey: crypto.publicEncrypt({ key: publicKey, oaepHash: "sha256" }, encryptionKey).toString("base64"),
  iv: iv.toString("base64"),
  tag: cipher.getAuthTag().toString("base64"),
  ciphertext: ciphertext.toString("base64")
};

fs.writeFileSync("/tmp/email-authorization-token.json", JSON.stringify(payload));
console.log("Microsoft authorization completed. The encrypted token artifact is ready.");

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
