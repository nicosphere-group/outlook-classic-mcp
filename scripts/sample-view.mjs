// Writes the confirmation window's contents for a made-up email, in Japanese, so the
// Windows checks can render the real window to a picture. All names are fictional.

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { buildView } from "../src/dialogText.js";
import { computeWarnings } from "../src/safety.js";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const snapshot = {
  sender_name: "山田 花子",
  sender_address: "yamada@example.com",
  account_type: "exchange",
  to: [{ name: "田中 太郎", address: "tanaka@example.co.jp" }],
  cc: [],
  bcc: [],
  subject: "Re: お見積りの件",
  body:
    "田中様\n\nお世話になっております。\n先日ご依頼いただいたお見積りを添付いたします。\n" +
    "ご不明な点がございましたら、お気軽にお知らせください。\n\nどうぞよろしくお願いいたします。\n\n山田",
  attachments: [{ index: 1, filename: "見積書.pdf", size: 123_000 }],
};
fs.mkdirSync(path.join(root, "dist"), { recursive: true });
for (const lang of ["ja", "en"]) {
  const view = buildView(snapshot, computeWarnings(snapshot), lang);
  fs.writeFileSync(path.join(root, "dist", `sample-view-${lang}.json`), JSON.stringify(view), "utf8");
}
