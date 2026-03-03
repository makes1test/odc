const ALLOWED_DOMAINS = new Set<string>(["es1.de"]);
const ALLOW_SUBDOMAINS = false;

type DomainCheckResult =
  | { ok: true; message?: never }
  | { ok: false; message: string };

type SmartEventCompletedOptions = {
  allowEvent: boolean;
  errorMessage?: string;
};

function extractDomain(email: string): string | null {
  const at = email.lastIndexOf("@");
  if (at < 0) return null;

  let domain = email.slice(at + 1).trim().toLowerCase();
  domain = domain.replace(/\.+$/, "");
  return domain || null;
}

function isAllowedDomain(domain: string): boolean {
  if (ALLOWED_DOMAINS.has(domain)) return true;
  if (!ALLOW_SUBDOMAINS) return false;

  for (const base of ALLOWED_DOMAINS) {
    if (domain.endsWith("." + base)) return true;
  }
  return false;
}

function getRecipientsAsync(field: Office.Recipients): Promise<any[]> {
  return new Promise((resolve, reject) => {
    try {
      field.getAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Failed) {
          reject(res.error);
          return;
        }
        resolve(res.value ?? []);
      });
    } catch (e) {
      reject(e);
    }
  });
}

async function checkDomains(): Promise<DomainCheckResult> {
  const item = Office.context.mailbox.item as any;

  const [to, cc, bcc] = await Promise.all([
    item?.to ? getRecipientsAsync(item.to) : Promise.resolve([]),
    item?.cc ? getRecipientsAsync(item.cc) : Promise.resolve([]),
    item?.bcc ? getRecipientsAsync(item.bcc) : Promise.resolve([]),
  ]);

  const emails: string[] = ([] as any[])
    .concat(to, cc, bcc)
    .map((r: any) => (typeof r?.emailAddress === "string" ? r.emailAddress.trim() : ""))
    .filter((x: string) => x.length > 0 && x.includes("@"));

  const domains = emails.map(extractDomain).filter((d): d is string => !!d);
  const uniqueDomains = Array.from(new Set(domains));

  if (uniqueDomains.length === 0) return { ok: true };

  const notAllowed = uniqueDomains.filter((d) => !isAllowedDomain(d));
  if (notAllowed.length > 0) {
    return {
      ok: false,
      message: `Nicht erlaubte Domain(s): ${notAllowed.join(", ")}.`,
    };
  }

  return { ok: true };
}

async function onMessageSendHandler(event: Office.AddinCommands.Event): Promise<void> {
  const FAILSAFE_MS = 5000;

  let finished = false;
  const done = (args: SmartEventCompletedOptions) => {
    if (finished) return;
    finished = true;
    event.completed(args as any);
  };

  const timer = setTimeout(() => {
    done({
      allowEvent: false,
      errorMessage: "Senden nicht möglich: Add-in Fehler bei Domainprüfung.",
    });
  }, FAILSAFE_MS);

  try {
    const result = await checkDomains();

    if (!result.ok) {
      done({
        allowEvent: false,
        errorMessage: result.message,
        // sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
      });
      return;
    }

    done({ allowEvent: true });
  } catch {
    done({
      allowEvent: false,
      errorMessage: "Senden nicht möglich: Add-in Fehler bei Domainprüfung.",
    });
  } finally {
    clearTimeout(timer);
  }
}

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});