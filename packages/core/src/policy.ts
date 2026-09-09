export type Policy = {
  allowedEndpoints?: string[];
  allowedModels?: string[];
  allowedSkills?: string[] | "*";
  denyExecuteJs: boolean;
  requireCompanion?: boolean;
  audit?: { enabled: boolean; redact: "none" | "prompts" | "values" };
  catalogUrl?: string;
};

export const OPEN_POLICY: Policy = {
  denyExecuteJs: true,
  audit: { enabled: false, redact: "prompts" }
};

/** User overrides may only narrow allowlists, never widen them. */
export function mergePolicy(tenant: Policy, user?: Partial<Policy>): Policy {
  const merged: Policy = { ...tenant, ...user, denyExecuteJs: tenant.denyExecuteJs || Boolean(user?.denyExecuteJs) };
  if (tenant.allowedEndpoints?.length) {
    const extra = user?.allowedEndpoints;
    const intersected = extra?.length
      ? tenant.allowedEndpoints.filter((e) => extra.some((u) => u.startsWith(e) || e.startsWith(u)))
      : tenant.allowedEndpoints;
    merged.allowedEndpoints = intersected.length ? intersected : tenant.allowedEndpoints;
  }
  if (tenant.allowedModels?.length) {
    const extra = user?.allowedModels;
    const intersected = extra?.length
      ? tenant.allowedModels.filter((m) => extra.includes(m))
      : tenant.allowedModels;
    merged.allowedModels = intersected.length ? intersected : tenant.allowedModels;
  }
  if (tenant.allowedSkills && tenant.allowedSkills !== "*") {
    const extra = user?.allowedSkills;
    if (extra && extra !== "*") {
      merged.allowedSkills = tenant.allowedSkills.filter((s) => extra.includes(s));
    } else {
      merged.allowedSkills = tenant.allowedSkills;
    }
  }
  if (tenant.requireCompanion) merged.requireCompanion = true;
  if (tenant.catalogUrl) merged.catalogUrl = tenant.catalogUrl;
  return merged;
}

export function assertPolicy(
  policy: Policy,
  request: { baseUrl: string; model: string; skill?: string; executeJs?: boolean }
): void {
  if (request.executeJs && policy.denyExecuteJs) {
    throw new Error("host.executeOfficeJs is disabled by policy.");
  }
  if (policy.allowedEndpoints && policy.allowedEndpoints.length > 0) {
    const ok = policy.allowedEndpoints.some((origin) => request.baseUrl.startsWith(origin));
    if (!ok) throw new Error("Endpoint is not allowed by policy.");
  }
  if (policy.allowedModels && policy.allowedModels.length > 0) {
    if (!policy.allowedModels.includes(request.model)) {
      throw new Error("Model is not allowed by policy.");
    }
  }
  if (request.skill && policy.allowedSkills && policy.allowedSkills !== "*") {
    if (!policy.allowedSkills.includes(request.skill)) {
      throw new Error("Skill is not allowed by policy.");
    }
  }
}
