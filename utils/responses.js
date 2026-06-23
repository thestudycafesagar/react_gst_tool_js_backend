/** Builds a TallyPushResponse-shaped object with the same field defaults
 * Pydantic always serializes (created/altered/ignored/errors=0, line_errors=[],
 * error=null) even when a given return site only sets a subset of them. */
function pushResponse(overrides = {}) {
  return {
    success: false,
    created: 0,
    altered: 0,
    ignored: 0,
    errors: 0,
    line_errors: [],
    error: null,
    ...overrides,
  };
}

module.exports = { pushResponse };
