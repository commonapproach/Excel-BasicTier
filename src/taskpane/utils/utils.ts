/* global document */

export function downloadJSONLD(data: any, filename: string): void {
  const jsonLDString = JSON.stringify(data, null, 2);
  const blob = new Blob([jsonLDString], { type: "application/ld+json" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}
