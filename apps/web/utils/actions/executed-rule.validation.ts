import { z } from "zod";

export const reviewExecutedRuleBody = z.object({
  executedRuleId: z.string(),
});
