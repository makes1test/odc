const ALLOWED_DOMAINS = new Set<string>(["es1.de"]);
const ALLOW_SUBDOMAINS = false;
type DomainCheckResult =
  | { ok: true; message?: never }
  | { ok: false; message: string };

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

  const bases = Array.from(ALLOWED_DOMAINS);
  for (let i = 0; i < bases.length; i++) {
    if (domain.endsWith("." + bases[i])) return true;
  }
  return false;
}

function getRecipientsAsync(
  field: Office.Recipients
): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve, reject) => {
    field.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        reject(res.error);
        return;
      }
      resolve(res.value ?? []);
    });
  });
}

async function checkDomains(): Promise<DomainCheckResult> {
  const item = Office.context.mailbox.item as any;

  const [to, cc, bcc] = await Promise.all([
    item?.to ? getRecipientsAsync(item.to) : Promise.resolve([]),
    item?.cc ? getRecipientsAsync(item.cc) : Promise.resolve([]),
    item?.bcc ? getRecipientsAsync(item.bcc) : Promise.resolve([]),
  ]);

  const emails: string[] = ([] as any[]).concat(to, cc, bcc)
    .map((r: any) => (typeof r?.emailAddress === "string" ? r.emailAddress.trim() : ""))
    .filter((x: string) => x.length > 0 && x.includes("@"));

  const domains = emails
    .map(extractDomain)
    .filter((d): d is string => !!d);

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

async function onMessageSendHandler(event: any) {
  try {
    const result: DomainCheckResult = await checkDomains();

    if (result.ok === false) {
      event.completed({
        allowEvent: false,
        errorMessage: result.message,
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser, 
      });
      return;
    }

    event.completed({ allowEvent: true });
  } catch (e: any) {
  event.completed({ allowEvent: true });
}
}

console.log("ODC: launchevent loaded");
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

Office.onReady(() => {
  console.log("ODC: Office.onReady");
});