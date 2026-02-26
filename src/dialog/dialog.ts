/* global Office */

/*const ALLOWED_DOMAINS = new Set<string>(["es1.de"]);

function extractDomain(email: string): string | null {
  const at = email.lastIndexOf("@");
  if (at < 0) return null;
  return email.slice(at + 1).trim().toLowerCase();
}

function getRecipientsAsync(field: Office.Recipients): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve, reject) => {
    field.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Failed) reject(res.error);
      else resolve(res.value || []);
    });
  });
}

function setStatus(text: string, kind: "ok" | "err" | "warn" = "ok") {
  const el = document.getElementById("msg");
  if (!el) return;
  el.textContent = text;
  el.classList.remove("ok", "err", "warn");
  el.classList.add(kind);
}

async function runDomainCheck() {
  const item = Office.context.mailbox.item as any;

  if (!item?.to || !item?.cc || !item?.bcc) {
    setStatus("Compose-Kontext nicht verfügbar (to/cc/bcc fehlen).", "err");
    return;
  }

  const [to, cc, bcc] = await Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
    getRecipientsAsync(item.bcc),
  ]);

  const emails = [...to, ...cc, ...bcc]
    .map((r: any) => r?.emailAddress)
    .filter((x: any) => typeof x === "string" && x.includes("@"));

  const domains = emails.map(extractDomain).filter((d): d is string => !!d);
  const uniqueDomains = Array.from(new Set(domains));

  if (uniqueDomains.length === 0) {
    setStatus("Keine Empfänger gefunden (To/CC/BCC leer).", "warn");
    return;
  }

  const notAllowed = uniqueDomains.filter((d) => !ALLOWED_DOMAINS.has(d));

  if (notAllowed.length > 0) {
    setStatus(`Nicht erlaubte Domain(s): ${notAllowed.join(", ")}`, "err");
    return;
  }

  setStatus("OK: Alle Empfänger-Domains sind erlaubt.", "ok");
}

Office.onReady(() => {
  const safeCheck = () =>
    runDomainCheck().catch((e) => setStatus("Fehler: " + (e?.message ?? String(e)), "err"));

  safeCheck();


  document.getElementById("recheck")?.addEventListener("click", safeCheck);

  const item = Office.context.mailbox.item as any;


  let t: any = null;
  const trigger = () => {
    if (t) clearTimeout(t);
    t = setTimeout(safeCheck, 200);
  };


  const canEvents = !!item?.to?.addHandlerAsync && !!Office?.EventType?.RecipientsChanged;
  if (canEvents) {
    item.to.addHandlerAsync(Office.EventType.RecipientsChanged, trigger);
    item.cc.addHandlerAsync(Office.EventType.RecipientsChanged, trigger);
    item.bcc.addHandlerAsync(Office.EventType.RecipientsChanged, trigger);
  }


  let lastKey = "";
  setInterval(async () => {
    try {
      if (!item?.to || !item?.cc || !item?.bcc) return;

      const [to, cc, bcc] = await Promise.all([
        getRecipientsAsync(item.to),
        getRecipientsAsync(item.cc),
        getRecipientsAsync(item.bcc),
      ]);

      const key = [...to, ...cc, ...bcc]
        .map((r: any) => r?.emailAddress || "")
        .join("|");

      if (key !== lastKey) {
        lastKey = key;
        trigger();
      }
    } catch {
    }
  }, 800);
}); */