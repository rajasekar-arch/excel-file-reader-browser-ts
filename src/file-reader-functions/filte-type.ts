// find file extension
export function getFileType(fileName: string): string {
  if (!fileName) {
    return "unknown";
  }
  const parts = fileName.split(".");
  if (parts.length < 2) {
    return "unknown";
  }

  return parts[parts.length - 1].toLowerCase();
}
