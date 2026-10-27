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
  } catch (error) {
    // The move or delete already happened, so there is nothing to undo.
    logger.warn("Could not record that a folder held mail", { error });
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
