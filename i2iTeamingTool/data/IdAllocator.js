/**
 * IdAllocator class for project ID generation.
 * Generates IDs in the format DIST-yy_yy-#### (e.g., NUSD-25_26-0024).
 * 
 * Note: Callers (processNewProjects, handleFormSubmission) are responsible for
 * acquiring a script lock before calling next(). This avoids nested lock conflicts
 * since Apps Script locks are not reentrant.
 */
class IdAllocator {
  /**
   * Creates a new IdAllocator instance.
   * @param {Config} config - The Config instance for reading district/year and serial
   */
  constructor(config) {
    this.config = config;
  }

  /**
   * Generates the next project ID.
   * IMPORTANT: Caller must hold a script lock before calling this method.
   * Format: DIST-yy_yy-#### (e.g., NUSD-25_26-0024)
   * @returns {string} The generated project ID
   * @throws {Error} If config is invalid
   */
  next() {
    const districtId = this.config.districtId;
    const schoolYear = this.config.schoolYear;

    if (!districtId) {
      throw new Error('IdAllocator: District ID is not configured');
    }

    if (!schoolYear) {
      throw new Error('IdAllocator: School Year is not configured');
    }

    // Get and increment the serial number
    const serial = this.config.getAndIncrementSerial();

    // Format the ID
    const formattedSerial = padNumber(serial, 4);
    const projectId = `${districtId}-${schoolYear}-${formattedSerial}`;

    DEBUG && console.log(`IdAllocator: Generated project ID: ${projectId}`);

    return projectId;
  }

  /**
   * Validates that a project ID matches the expected format.
   * @param {string} projectId - The project ID to validate
   * @returns {boolean} True if the format is valid
   */
  isValidFormat(projectId) {
    if (!projectId || typeof projectId !== 'string') {
      return false;
    }

    // Expected format: DIST-yy_yy-####
    // Example: NUSD-25_26-0024
    const pattern = /^[A-Z]{2,10}-\d{2}_\d{2}-\d{4}$/;
    return pattern.test(projectId);
  }

  /**
   * Parses a project ID into its components.
   * @param {string} projectId - The project ID to parse
   * @returns {Object|null} Object with districtId, schoolYear, serial, or null if invalid
   */
  parse(projectId) {
    if (!this.isValidFormat(projectId)) {
      return null;
    }

    const parts = projectId.split('-');
    return {
      districtId: parts[0],
      schoolYear: parts[1],
      serial: parseInt(parts[2], 10)
    };
  }

  /**
   * Gets the current serial number without incrementing.
   * @returns {number} The current next serial value
   */
  peekNextSerial() {
    return this.config.nextSerial;
  }
}

