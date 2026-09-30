import { describe, expect, it } from "vitest";
import { format, hostLocales, isRtl, localize, setLocale } from "./control-ui-i18n.js";

describe("Microsoft Graph Control UI localization", () => {
  it("uses the host locale and English fallback for untranslated host locales", () => {
    setLocale("de-DE");
    expect(localize("Dienste")).toBe("Dienste");
    setLocale("es");
    expect(localize("Dienste")).toBe("Servicios");
    setLocale("fr");
    expect(localize("Dienste")).toBe("Services");
    expect(hostLocales).toHaveLength(21);
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
