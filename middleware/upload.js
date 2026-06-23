/** Memory-storage multer instance — mirrors FastAPI's UploadFile.read(),
 * which reads the whole upload into memory (no disk persistence). */
const multer = require('multer');

module.exports = multer({ storage: multer.memoryStorage() });
