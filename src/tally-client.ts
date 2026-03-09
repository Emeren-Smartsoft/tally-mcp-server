/**
 * Tally HTTP Client
 * Sends XML requests to TallyPrime's HTTP XML API and returns raw XML responses.
 */

import axios, { AxiosError } from 'axios';

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const TALLY_HOST = process.env.TALLY_XML_HOST ?? 'localhost';
const TALLY_PORT = process.env.TALLY_XML_PORT ?? '9000';
const TALLY_BASE_URL = `http://${TALLY_HOST}:${TALLY_PORT}`;

/** Timeout in milliseconds — 60 s gives Tally time for large queries. */
const REQUEST_TIMEOUT_MS = 60_000;

// ---------------------------------------------------------------------------
// Client
// ---------------------------------------------------------------------------

/**
 * Send an XML request to TallyPrime and return the raw XML response string.
 *
 * @throws {Error} if TallyPrime is unreachable or returns a non-2xx status.
 */
export async function sendTallyRequest(xmlBody: string): Promise<string> {
  try {
    const response = await axios.post<string>(TALLY_BASE_URL, xmlBody, {
      headers: {
        'Content-Type': 'application/xml; charset=utf-8',
      },
      timeout: REQUEST_TIMEOUT_MS,
      responseType: 'text',
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
 * Check whether TallyPrime is reachable.
 * Returns true if a basic ping succeeds, false otherwise.
 */
export async function isTallyReachable(): Promise<boolean> {
  try {
    await axios.get(TALLY_BASE_URL, { timeout: 5_000 });
    return true;
  } catch (err) {
    const axiosErr = err as AxiosError;
    // A 4xx / 5xx response still means Tally is up; ECONNREFUSED means it isn't.
    if (axiosErr.response) return true;
    return false;
  }
}
