// Errors whose messages are shown to the model and, through it, to the user.
// Every message is written for a non-technical reader.

export class OutlookMcpError extends Error {
  constructor(message) {
    super(message);
    this.name = this.constructor.name;
  }
}

/** Outlook could not be reached (not installed, not responding, elevation mismatch). */
export class OutlookUnavailableError extends OutlookMcpError {}

/** A mail, folder, account or attachment does not exist (any more). */
export class NotFoundError extends OutlookMcpError {}

/** The item exists but is not an unsent draft. */
export class NotADraftError extends OutlookMcpError {}

/** Outside what this tool permits (path limits, disabled send). */
export class NotAllowedError extends OutlookMcpError {}

/** An argument could not be understood. */
export class InvalidInputError extends OutlookMcpError {}
