//JSON.stringify(data, null, 2)
export function downloadJsonAsFile(data: string, filename = "chats.json") {
  const blob = new Blob(
    [data],
    { type: "application/json" }
  );

  downloadBlobAsFile(blob, filename);
}

export function downloadBlobAsFile(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";

  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);    
}