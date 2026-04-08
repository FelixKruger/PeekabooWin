export class CliUsageError extends Error {
  constructor(message) {
    super(message);
    this.name = "CliUsageError";
  }
}

export class BackendError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = "BackendError";
    this.details = details;
  }
}
