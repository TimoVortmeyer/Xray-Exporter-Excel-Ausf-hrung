(async () => {
  const sharePointContext = window._spPageContextInfo;
  if (!sharePointContext?.webServerRelativeUrl || !sharePointContext?.webAbsoluteUrl) {
    throw new Error("SharePoint-Seitenkontext nicht gefunden. Bitte auf einer SharePoint-Seite starten.");
  }

  const siteOrigin = window.location.origin;
  const webUrl = new URL(sharePointContext.webAbsoluteUrl);
  const sitePath = sharePointContext.webServerRelativeUrl.replace(/\/$/, "");
  const allowedApiPath = `${sitePath}/_api/`;
  const currentUrl = new URL(window.location.href);
  const selectedFolder = currentUrl.searchParams.get("id") || currentUrl.searchParams.get("RootFolder");
  const requestedFolder = selectedFolder || window.prompt(
    "Server-relative Pfad des zu durchsuchenden SharePoint-Ordners eingeben:",
    `${sitePath}/Shared Documents`
  );

  if (!requestedFolder) {
    throw new Error("Kein SharePoint-Ordner ausgewaehlt.");
  }
  if (window.location.protocol !== "https:" || webUrl.origin !== siteOrigin) {
    throw new Error("Der SharePoint-Seitenkontext passt nicht zur aktuellen HTTPS-Seite.");
  }

  const requestedFolderUrl = new URL(requestedFolder, siteOrigin);
  const startFolder = decodeURIComponent(requestedFolderUrl.pathname).replace(/\/$/, "");
  const normalizedSitePath = sitePath.toLowerCase();
  const normalizedStartFolder = startFolder.toLowerCase();
  if (requestedFolderUrl.origin !== siteOrigin ||
      (normalizedStartFolder !== normalizedSitePath &&
       !normalizedStartFolder.startsWith(`${normalizedSitePath}/`))) {
    throw new Error("Der ausgewaehlte Ordner liegt ausserhalb der aktuellen SharePoint-Site.");
  }

  const files = [];
  let folderCount = 0;

  function encodeODataString(value) {
    return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, "%27");
  }

  function folderEndpoint(path) {
    return (
      `${siteOrigin}${sitePath}/_api/web/` +
      `GetFolderByServerRelativePath(decodedurl='${encodeODataString(path)}')`
    );
  }

  function validateApiUrl(value, baseUrl) {
    const url = new URL(value, baseUrl);
    if (url.origin !== siteOrigin ||
        !url.pathname.toLowerCase().startsWith(allowedApiPath.toLowerCase())) {
      throw new Error("SharePoint hat einen unerwarteten API-Folgelink geliefert.");
    }
    return url.href;
  }

  function fileLink(path) {
    const url = new URL(path, siteOrigin);
    const normalizedPath = url.pathname.toLowerCase();
    if (url.origin !== siteOrigin ||
      (normalizedPath !== normalizedSitePath &&
       !normalizedPath.startsWith(`${normalizedSitePath}/`))) {
      throw new Error("SharePoint hat einen Dateilink ausserhalb der erwarteten Site geliefert.");
    }
    return url.href;
  }

  async function getAll(url) {
    const results = [];
    let nextUrl = validateApiUrl(url, siteOrigin);

    while (nextUrl) {
      const response = await fetch(nextUrl, {
        credentials: "same-origin",
        headers: { Accept: "application/json;odata=nometadata" }
      });

      if (!response.ok) {
        throw new Error(
          `SharePoint-Anfrage fehlgeschlagen: ${response.status} ${response.statusText}\n${nextUrl}`
        );
      }

      const data = await response.json();
      results.push(...(data.value ?? data.d?.results ?? []));
      const next = data["@odata.nextLink"] ?? data["odata.nextLink"] ?? data.d?.__next;
      nextUrl = next ? validateApiUrl(next, nextUrl) : null;
    }

    return results;
  }

  const pendingFolders = [startFolder];

  while (pendingFolders.length > 0) {
    const folderPath = pendingFolders.pop();
    const endpoint = folderEndpoint(folderPath);
    const currentFiles = await getAll(endpoint + "/Files?$select=Name,ServerRelativeUrl");

    for (const file of currentFiles) {
      files.push({
        Dateiname: file.Name,
        Ablagepfad: file.ServerRelativeUrl,
        Link: fileLink(file.ServerRelativeUrl)
      });
    }

    const subfolders = await getAll(endpoint + "/Folders?$select=Name,ServerRelativeUrl");
    for (const folder of subfolders) {
      pendingFolders.push(folder.ServerRelativeUrl);
    }

    folderCount++;
    console.log(`${folderCount} Ordner geprueft, ${files.length} Dateien gefunden`);
  }

  files.sort((left, right) => left.Ablagepfad.localeCompare(right.Ablagepfad, "de"));
  window.sharePointDateiliste = files;
  console.table(files);

  function csvCell(value) {
    return `"${String(value).replace(/"/g, '""')}"`;
  }

  const csv = [
    ["Dateiname", "Ablagepfad", "Link"],
    ...files.map(file => [file.Dateiname, file.Ablagepfad, file.Link])
  ]
    .map(row => row.map(csvCell).join(";"))
    .join("\r\n");

  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
  const buttonId = "xray-sharepoint-csv-save";
  document.getElementById(buttonId)?.remove();
  const saveButton = document.createElement("button");
  saveButton.id = buttonId;
  saveButton.type = "button";
  saveButton.textContent = `SharePoint-CSV speichern (${files.length} Dateien)`;
  Object.assign(saveButton.style, {
    position: "fixed",
    right: "20px",
    bottom: "20px",
    zIndex: "2147483647",
    padding: "12px 18px",
    color: "#ffffff",
    background: "#0050a4",
    border: "0",
    borderRadius: "4px",
    font: "600 14px Segoe UI, sans-serif",
    cursor: "pointer"
  });

  saveButton.addEventListener("click", async () => {
    saveButton.disabled = true;
    try {
      if (typeof window.showSaveFilePicker !== "function") {
        throw new Error("Direktes Speichern wird nicht unterstuetzt. Bitte Edge oder Chrome verwenden.");
      }
      const handle = await window.showSaveFilePicker({
        id: "XraySharePointDateiliste",
        suggestedName: "SharePoint-Dateiliste.csv",
        types: [{ description: "CSV-Datei", accept: { "text/csv": [".csv"] } }]
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      saveButton.textContent = "SharePoint-CSV gespeichert";
      console.log(`CSV direkt gespeichert: ${handle.name}`);
    } catch (error) {
      if (error.name === "AbortError") {
        console.log("Speichern abgebrochen.");
      } else {
        console.error("CSV konnte nicht gespeichert werden:", error);
      }
    } finally {
      saveButton.disabled = false;
    }
  });

  document.body.appendChild(saveButton);
  console.log(`Fertig: ${files.length} Dateien aus ${folderCount} Ordnern. Zum Speichern den SharePoint-CSV-Button anklicken.`);
})().catch(error => {
  console.error("Export abgebrochen:", error);
});
