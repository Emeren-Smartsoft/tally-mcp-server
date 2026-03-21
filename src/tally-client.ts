/**
 * Tally HTTP Client
 * Sends XML requests to TallyPrime's HTTP XML API and returns raw XML responses.
 */

import axios, { AxiosError } from 'axios';
import http from 'http';

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const TALLY_HOST = process.env.TALLY_XML_HOST ?? 'localhost';
const TALLY_PORT = process.env.TALLY_XML_PORT ?? '9000';
const TALLY_BASE_URL = `http://${TALLY_HOST}:${TALLY_PORT}`;

/** Timeout in milliseconds — 30 s is sufficient for most Tally queries. */
const REQUEST_TIMEOUT_MS = 30_000;

/** Maximum number of retry attempts for failed requests. */
const MAX_RETRIES = 3;

/** Maximum number of concurrent requests to TallyPrime. */
const MAX_CONCURRENT_REQUESTS = 2;

// ---------------------------------------------------------------------------
// HTTP Keep-Alive Agent
// ---------------------------------------------------------------------------

/** Reuse connections to reduce handshake overhead per request. */
const keepAliveAgent = new http.Agent({
  keepAlive: true,
  maxSockets: MAX_CONCURRENT_REQUESTS,
});

// ---------------------------------------------------------------------------
// Request Queue
// ---------------------------------------------------------------------------

/**
 * Simple queue that limits the number of in-flight requests to TallyPrime.
 * Prevents resource contention and avoids overwhelming the Tally process.
 */
class RequestQueue {
  private queue: Array<() => void> = [];
  private activeRequests = 0;

  async enqueue<T>(fn: () => Promise<T>): Promise<T> {
    await this.waitForSlot();
    this.activeRequests++;
    try {
      return await fn();
    } finally {
      this.activeRequests--;
      this.releaseSlot();
    }
  }

  private waitForSlot(): Promise<void> {
    if (this.activeRequests < MAX_CONCURRENT_REQUESTS) {
      return Promise.resolve();
    }
    return new Promise<void>((resolve) => {
      this.queue.push(resolve);
    });
  }

  private releaseSlot(): void {
    const next = this.queue.shift();
    if (next) next();
  }
}

const requestQueue = new RequestQueue();

// ---------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------

/**
 * Send an XML request to TallyPrime and return the raw XML response string.
 * Runs through the global request queue (max 2 concurrent) and does NOT retry.
 *
 * @throws {Error} if TallyPrime is unreachable or returns a non-2xx status.
 */
export async function sendTallyRequest(xmlBody: string): Promise<string> {
  return requestQueue.enqueue(() => sendOnce(xmlBody));
}

/**
 * Send a single HTTP request without queueing or retry logic.
 * Internal helper used by sendTallyRequest and sendTallyRequestWithRetry.
 */
async function sendOnce(xmlBody: string): Promise<string> {
  try {
    const response = await axios.post<string>(TALLY_BASE_URL, xmlBody, {
      headers: {
        'Content-Type': 'application/xml; charset=utf-8',
      },
      timeout: REQUEST_TIMEOUT_MS,
      responseType: 'text',
      httpAgent: keepAliveAgent,
    });
    return response.data;
  } catch (err) {
    const axiosErr = err as AxiosError;
    if (axiosErr.code === 'ECONNREFUSED' || axiosErr.code === 'ECONNRESET') {
      throw new Error(
        `Cannot connect to TallyPrime at ${TALLY_BASE_URL}. ` +
          'Ensure TallyPrime is running and a company is loaded.'
      );
    }
    if (axiosErr.code === 'ETIMEDOUT' || axiosErr.code === 'ECONNABORTED') {
      throw new Error(
        `Request to TallyPrime timed out after ${REQUEST_TIMEOUT_MS / 1000}s. ` +
          'The company may be large or TallyPrime may be busy.'
      );
    }
    const message =
      axiosErr.message ?? (err instanceof Error ? err.message : 'Unknown error contacting Tally');
    throw new Error(`TallyPrime request failed: ${message}`);
  }
}

/**
 * Send an XML request to TallyPrime with automatic retry on transient failures.
 * Uses exponential backoff between retries to avoid hammering a busy Tally.
 *
 * @param xmlBody  - The XML payload to send.
 * @param maxRetries - Maximum number of attempts (default: MAX_RETRIES = 3).
 * @throws {Error} after all retries are exhausted.
 */
export async function sendTallyRequestWithRetry(
  xmlBody: string,
  maxRetries: number = MAX_RETRIES
): Promise<string> {
  let lastError: Error = new Error('Unknown error');

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await requestQueue.enqueue(() => sendOnce(xmlBody));
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));

      // Do not retry on connection-refused — Tally is simply not running.
      if (lastError.message.includes('Cannot connect to TallyPrime')) {
        throw lastError;
      }

      if (attempt < maxRetries) {
        // Exponential backoff: 500 ms, 1000 ms, 2000 ms, ...
        const delayMs = 500 * Math.pow(2, attempt - 1);
        await new Promise<void>((resolve) => setTimeout(resolve, delayMs));
      }
    }
  }

  throw new Error(`TallyPrime request failed after ${maxRetries} attempts: ${lastError.message}`);
}

/**
 * Check whether TallyPrime is reachable.
 * Returns true if a basic ping succeeds, false otherwise.
 */
export async function isTallyReachable(): Promise<boolean> {
  try {
    await axios.get(TALLY_BASE_URL, { timeout: 5_000, httpAgent: keepAliveAgent });
    return true;
  } catch (err) {
    const axiosErr = err as AxiosError;
    // A 4xx / 5xx response still means Tally is up; ECONNREFUSED means it isn't.
    if (axiosErr.response) return true;
    return false;
  }
}
