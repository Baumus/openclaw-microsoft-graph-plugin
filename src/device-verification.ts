/** Only browser destinations returned by Microsoft's device-code endpoint. */
export function isMicrosoftDeviceVerificationUri(value: unknown): value is string {
  return value === "https://login.microsoft.com/device"
    || value === "https://microsoft.com/devicelogin"
    || value === "https://www.microsoft.com/devicelogin"
    || value === "https://www.microsoft.com/link";
}
