const FEATURE_POLICY = /** @type {const} */ ({
  "mail-read": {
    scopes: ["Mail.Read"],
    implies: [],
  },
  "mail-write": {
    scopes: ["Mail.ReadWrite"],
    implies: ["mail-read"],
  },
  "mail-send": {
    scopes: ["Mail.Send"],
    implies: [],
  },
  "calendar-read": {
    scopes: ["Calendars.Read"],
    implies: [],
  },
  "calendar-write": {
    scopes: ["Calendars.ReadWrite"],
    implies: ["calendar-read"],
  },
  "tasks-read": {
    scopes: ["Tasks.Read"],
    implies: [],
  },
  "tasks-write": {
    scopes: ["Tasks.ReadWrite"],
    implies: ["tasks-read"],
  },
  "onedrive-read": {
    scopes: ["Files.Read"],
    implies: [],
  },
  "onedrive-write": {
    scopes: ["Files.ReadWrite"],
    implies: ["onedrive-read"],
  },
});

for (const definition of Object.values(FEATURE_POLICY)) {
  Object.freeze(definition.scopes);
  Object.freeze(definition.implies);
  Object.freeze(definition);
}

export const M365_FEATURE_POLICY = Object.freeze(FEATURE_POLICY);

/** @typedef {keyof typeof M365_FEATURE_POLICY} M365Feature */

export const M365_FEATURES = Object.freeze(
  /** @type {M365Feature[]} */ (Object.keys(M365_FEATURE_POLICY)),
);

export const DEFAULT_M365_FEATURES = Object.freeze(
  /** @type {const} */ ([
    "calendar-write",
    "mail-write",
    "mail-send",
    "tasks-write",
  ]),
);

export class UnknownM365FeatureError extends Error {
  /** @param {readonly string[]} unknownFeatures */
  constructor(unknownFeatures) {
    super(`Unknown Microsoft 365 feature(s): ${unknownFeatures.join(", ")}`);
    this.name = "UnknownM365FeatureError";
    this.unknownFeatures = [...unknownFeatures];
  }
}

/**
 * @param {readonly unknown[]} values
 * @returns {M365Feature[]}
 */
export function normalizeM365Features(values) {
  /** @type {M365Feature[]} */
  const normalized = [];
  const seen = new Set();
  const unknown = [];

  for (const value of values) {
    if (typeof value !== "string") {
      throw new TypeError("Microsoft 365 features must be strings");
    }
    const feature = value.trim().toLowerCase();
    if (!feature) continue;
    if (!Object.prototype.hasOwnProperty.call(M365_FEATURE_POLICY, feature)) {
      unknown.push(feature);
      continue;
    }
    if (!seen.has(feature)) {
      seen.add(feature);
      normalized.push(/** @type {M365Feature} */ (feature));
    }
  }

  if (unknown.length) throw new UnknownM365FeatureError(unknown);
  return normalized;
}

/**
 * @param {string | readonly unknown[] | null | undefined} value
 * @param {readonly M365Feature[]} [fallback]
 * @returns {M365Feature[]}
 */
export function parseM365Features(value, fallback = DEFAULT_M365_FEATURES) {
  if (value === undefined || value === null) {
    return normalizeM365Features(fallback);
  }
  if (typeof value === "string") {
    return normalizeM365Features(value.split(","));
  }
  if (!Array.isArray(value)) {
    throw new TypeError("Microsoft 365 features must be an array or comma-separated string");
  }
  return normalizeM365Features(value);
}

/**
 * @param {readonly unknown[]} features
 * @returns {M365Feature[]}
 */
export function expandM365Features(features) {
  const expanded = normalizeM365Features(features);
  const seen = new Set(expanded);

  for (let index = 0; index < expanded.length; index += 1) {
    const feature = expanded[index];
    for (const implied of M365_FEATURE_POLICY[feature].implies) {
      if (!seen.has(implied)) {
        seen.add(implied);
        expanded.push(implied);
      }
    }
  }

  return expanded;
}

/**
 * @param {readonly unknown[]} features
 * @param {unknown} requiredFeature
 * @returns {boolean}
 */
export function isM365FeatureEnabled(features, requiredFeature) {
  const [required] = normalizeM365Features([requiredFeature]);
  return required !== undefined && expandM365Features(features).includes(required);
}

/**
 * @param {readonly unknown[]} features
 * @returns {string[]}
 */
export function deriveFeatureGraphScopes(features) {
  const scopes = [];
  const seen = new Set();
  for (const feature of normalizeM365Features(features)) {
    for (const scope of M365_FEATURE_POLICY[feature].scopes) {
      const key = scope.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        scopes.push(scope);
      }
    }
  }
  return scopes;
}

/**
 * @param {readonly unknown[]} features
 * @returns {string[]}
 */
export function deriveAllowedGraphScopes(features) {
  return deriveFeatureGraphScopes(expandM365Features(features));
}

/**
 * @param {readonly unknown[]} scopes
 * @returns {string[]}
 */
export function normalizeGraphScopeNames(scopes) {
  const normalized = [];
  const seen = new Set();
  for (const value of scopes) {
    if (typeof value !== "string") {
      throw new TypeError("Microsoft Graph scopes must be strings");
    }
    const scope = value.trim();
    if (!scope) continue;
    const key = scope.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      normalized.push(scope);
    }
  }
  return normalized;
}

/**
 * @param {readonly unknown[]} scopes
 * @returns {{ features: M365Feature[], unknownScopes: string[] }}
 */
export function inferM365FeaturesFromScopes(scopes) {
  /** @type {Map<string, M365Feature>} */
  const scopeToFeature = new Map();
  for (const feature of M365_FEATURES) {
    for (const scope of M365_FEATURE_POLICY[feature].scopes) {
      scopeToFeature.set(scope.toLowerCase(), feature);
    }
  }

  const features = [];
  const unknownScopes = [];
  for (const scope of normalizeGraphScopeNames(scopes)) {
    const feature = scopeToFeature.get(scope.toLowerCase());
    if (feature) features.push(feature);
    else unknownScopes.push(scope);
  }
  return {
    features: normalizeM365Features(features),
    unknownScopes,
  };
}

/**
 * @param {readonly unknown[]} requestedScopes
 * @param {readonly unknown[]} features
 * @returns {string[]}
 */
export function getDisallowedGraphScopes(requestedScopes, features) {
  const allowed = new Set(
    deriveAllowedGraphScopes(features).map((scope) => scope.toLowerCase()),
  );
  return normalizeGraphScopeNames(requestedScopes).filter(
    (scope) => !allowed.has(scope.toLowerCase()),
  );
}
