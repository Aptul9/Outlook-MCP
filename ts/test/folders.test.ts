import assert from "node:assert/strict";
import { test } from "node:test";
import { listFolders } from "../src/api.ts";
import { fake } from "./helpers.ts";
test("folder pagination and deeper children preserved; unsafe nextLink fails explicitly", async () => {
  const { o } = fake({
    "GET me/MailFolders?": (c) =>
      c.url.includes("skip")
        ? { value: [{ Id: "F2", DisplayName: "Second" }] }
        : {
            value: [{ Id: "F1", DisplayName: "First", ChildFolderCount: 1 }],
            "@odata.nextLink":
              "https://outlook.office.com/api/v2.0/me/MailFolders?$skip=100",
          },
    "GET me/MailFolders/F1/childfolders": () => ({
      value: [{ Id: "F3", DisplayName: "Child", ChildFolderCount: 1 }],
    }),
    "GET me/MailFolders/F3/childfolders": () => ({
      value: [{ Id: "F4", DisplayName: "Deep" }],
    }),
  });
  assert.deepEqual(
    (await listFolders(o)).map((f) => f.name),
    ["First", "First/Child", "First/Child/Deep", "Second"],
  );
  const bad = fake({
    "GET me/MailFolders": () => ({
      value: [],
      "@odata.nextLink": "https://evil.example/x",
    }),
  });
  await assert.rejects(listFolders(bad.o), /Untrusted/);
});
