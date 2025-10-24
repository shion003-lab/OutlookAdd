Office.onReady(() => {
  document.getElementById("saveBtn").addEventListener("click", async () => {
    const assignee = document.getElementById("assignee").value;
    const status = document.getElementById("status").value;
    const notes = document.getElementById("notes").value;
    const result = document.getElementById("result");

    try {
      const item = Office.context.mailbox.item;
      const subject = item.subject;
      const sender = item.from.displayName || item.from.emailAddress;
      const id = item.itemId;

      result.textContent = `保存テスト: 件名="${subject}" 担当=${assignee}, ステータス=${status}`;
    } catch (e) {
      console.error(e);
      result.textContent = "エラーが発生しました: " + e.message;
    }
  });
});
