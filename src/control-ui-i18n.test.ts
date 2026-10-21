import { describe, expect, it } from "vitest";
import { format, hostLocales, isRtl, localize, setLocale } from "./control-ui-i18n.js";

describe("Microsoft Graph Control UI localization", () => {
  it("uses the host locale and English fallback for untranslated host locales", () => {
    setLocale("de-DE");
    expect(localize("Dienste")).toBe("Dienste");
    expect(localize("Vault-SecretRef vorhanden")).toBe("Vault-SecretRef vorhanden");
    setLocale("en");
    expect(localize("Vault-SecretRef vorhanden")).toBe("Vault SecretRef present");
    setLocale("es");
    expect(localize("Dienste")).toBe("Servicios");
    setLocale("fr");
    expect(localize("Dienste")).toBe("Services");
    expect(hostLocales).toHaveLength(21);
  });

  it("translates the browser sign-in actions in all supported locales", () => {
    for (const code of ["en", "de", "es", "ar"]) {
      setLocale(code);
      expect(localize("Mit Microsoft verbinden")).toBeTruthy();
      expect(localize("Microsoft-Anmeldung öffnen")).toBeTruthy();
      expect(localize("Anmeldung abbrechen")).toBeTruthy();
      if (code !== "de") expect(localize("Microsoft-Anmeldung öffnen")).not.toBe("Microsoft-Anmeldung öffnen");
    }
  });

  it("translates the reauthorization guidance in all supported locales", () => {
    for (const code of ["en", "de", "es", "ar"]) {
      setLocale(code);
      expect(localize("Microsoft-Zugang erneut autorisieren")).toBeTruthy();
      expect(format("Diese Berechtigungen fehlen: {scopes}. Bis zur Erneuerung bleiben die betreffenden Aktionen gesperrt.", { scopes: "Contacts.ReadWrite" })).toContain("Contacts.ReadWrite");
      if (code !== "de") expect(localize("Berechtigungen erneut prüfen")).not.toBe("Berechtigungen erneut prüfen");
      expect(localize("Anwendungs-ID (Client-ID)")).toContain(code === "de" ? "Client-ID" : "client");
      expect(localize("Die vorhandene App und der vorhandene Mandant werden sicher wiederverwendet. Du musst keine Kennungen eingeben und keine neue App anlegen.")).not.toContain("terminal");
    }
  });

  it("localizes compact connection copy without claiming a tested read", () => {
    for (const code of ["en", "de", "es", "ar"]) {
      setLocale(code);
      expect(localize("Verbindung hergestellt")).not.toBe("");
      expect(localize("Update verfügbar")).not.toBe("");
      expect(localize("Update verfügbar: Version {version}")).not.toBe("");
      expect(localize("Verbindung hergestellt. Ein Lesezugriff durch einen berechtigten Agenten wurde nicht geprüft.")).not.toBe("");
      if (code !== "de") expect(localize("Verbindung hergestellt")).not.toBe("Verbindung hergestellt");
    }
  });

  it("preserves dynamic values and handles RTL locales", () => {
    setLocale("ar");
    expect(isRtl()).toBe(true);
    expect(format("Den Ordner {path} für alle Agenten entfernen?", { path: "/Clients/{x}" })).toContain("/Clients/{x}");
    setLocale("fa");
    expect(isRtl()).toBe(false);
    setLocale("en");
    expect(isRtl()).toBe(false);
  });
});
