import multer from "multer";
import path from "path";

// for upload excel
const storageExcel = multer.memoryStorage({
  destination: function (req, file, cb) {
    return cb(null, "./public/import_excel");
  },
  filename: function (req, file, cb) {
    return cb(
      null,
      `${Date.now()}-${Math.floor(1000 + Math.random() * 9000)}${
        path.parse(file.originalname).ext
      }`
    );
  },
});

const uploadExcel = multer({
  storage: storageExcel,
  fileFilter: (req, file, cb) => {
    // console.log(file);
    if (
      file.mimetype ==
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) {
      cb(null, true);
    } else {
      cb(null, false);
      return cb(new Error("Invalid file type. Only Excel files are allowed."));
    }
  },
});



export {uploadExcel };
