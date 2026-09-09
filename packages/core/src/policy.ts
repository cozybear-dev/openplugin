export type Policy = {
  allowedEndpoints?: string[];
  allowedModels?: string[];
  allowedSkills?: string[] | "*";
  denyExecuteJs: boolean;
  requireCompanion?: boolean;
};

export const OPEN_POLICY: Policy = {
  denyExecuteJs: true
};

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
