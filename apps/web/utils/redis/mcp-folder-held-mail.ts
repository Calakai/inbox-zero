import "server-only";
import { createHash } from "node:crypto";
import { createRateLimitKey } from "@/utils/rate-limit";
import { redis } from "@/utils/redis";
import type { Logger } from "@/utils/logger";

const HELD_MAIL_TTL_SECONDS = 24 * 60 * 60;

export async function markFolderHeldMail({
  emailAccountId,
  folderId,
  logger,
}: {
  emailAccountId: string;
  folderId: string;
  logger: Logger;
}) {
  try {
    await redis.set(getKey({ emailAccountId, folderId }), "1", {
      ex: HELD_MAIL_TTL_SECONDS,
    });
    return true;
  } catch (error) {
    // Callers record before moving or deleting and stop on false, so a lost
    // record can never leave an emptied folder deletable without asking.
    logger.warn("Could not record that a folder held mail", { error });
    return false;
  }
}

export async function wasFolderRecentlyHeldMail({
  emailAccountId,
  folderId,
  logger,
}: {
  emailAccountId: string;
  folderId: string;
  logger: Logger;
}) {
  try {
    return Boolean(await redis.get(getKey({ emailAccountId, folderId })));
  } catch (error) {
    // Fail closed: without the record we cannot tell an emptied folder from an
    // untouched one, and a wrong "untouched" would delete without asking.
    logger.warn("Could not check whether a folder held mail; asking", {
      error,
    });
    return true;
  }
}

function getKey({
  emailAccountId,
  folderId,
}: {
  emailAccountId: string;
  folderId: string;
}) {
  return createRateLimitKey([
    "mcp",
    "folder-held-mail",
    emailAccountId,
    createHash("sha256").update(folderId).digest("hex"),
  ]);
}
