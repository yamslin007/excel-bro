import capabilities from "../../../config/capabilities.json";

export type WpsWriteVersionStatus = "verified" | "development" | "blocked";

function matchesVersionPrefix(version: string, prefix: string): boolean {
  return version === prefix || version.startsWith(`${prefix}.`);
}

/**
 * Application.Version is the JSAPI-reported ET version, not the installation
 * directory's product version. The observed "12.0" is admitted only while
 * running the development add-in so real-host write/undo can be validated.
 */
export function wpsWriteVersionStatus(
  hostVersion: string | null | undefined,
  development = import.meta.env.DEV
): WpsWriteVersionStatus {
  const version = (hostVersion ?? "").trim();
  if (!version) return "blocked";
  if (capabilities.wpsWrite.verifiedVersionPrefixes.some((prefix) =>
    matchesVersionPrefix(version, prefix)
  )) return "verified";
  if (development && capabilities.wpsWrite.developmentVersionPrefixes.some((prefix) =>
    matchesVersionPrefix(version, prefix)
  )) return "development";
  return "blocked";
}
