import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import path from "path";
import dotenv from "dotenv";
import readXlsxFile from "read-excel-file/node";
import { Client } from "@neondatabase/serverless";
import ExcelJS from "exceljs";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 8000;

app.use(cors({
  origin: ["http://localhost:3000"],
  methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const uploadPath = path.join(process.cwd(), "uploads");
if (!fs.existsSync(uploadPath)) fs.mkdirSync(uploadPath, { recursive: true });
app.use("/uploads", express.static(uploadPath));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadPath),
  filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname),
});
const upload = multer({ storage });

const getClient = () => new Client({
  connectionString: process.env.NEON_URL,
  ssl: { rejectUnauthorized: false }
});

app.get("/", (req, res) => res.send("Backend is running!"));

// --- Upload Excel ---
app.post("/upload", upload.single("excel"), async (req, res) => {
  if (!req.file) return res.status(400).send("No file uploaded");

  let { tableName } = req.body;
  const rows = await readXlsxFile(req.file.path);
  if (rows.length < 2) return res.status(400).send("Excel has no data");

  const headers = rows[0];
  const dataRows = rows.slice(1);

  if (!tableName) tableName = path.parse(req.file.originalname).name;

  const client = getClient();
  await client.connect();

  try {
    // Ensure ID column exists
    const hasID = headers.includes("ID");
    if (!hasID) headers.unshift("ID"); // add ID column
    const columnsSQL = headers.map(h => `"${h}" TEXT`).join(", ");
    const createSQL = `CREATE TABLE IF NOT EXISTS "${tableName}" (${columnsSQL})`;
    await client.query(createSQL);

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const values = hasID ? row : [i + 1, ...row]; // add ID if missing
      const placeholders = values.map((_, idx) => `$${idx + 1}`).join(", ");
      const insertSQL = `INSERT INTO "${tableName}" (${headers.map(h => `"${h}"`).join(", ")}) VALUES (${placeholders})`;
      await client.query(insertSQL, values);
    }

    res.json({
      message: "Excel inserted successfully",
      table: tableName,
      inserted: dataRows.length
    });
  } catch (err) {
    console.error("Upload Error:", err);
    res.status(500).send("Error: " + err.message);
  } finally { await client.end(); }
});

// --- Get all tables ---
app.get("/tables", async (req, res) => {
  const client = getClient();
  await client.connect();
  try {
    const result = await client.query(`
      SELECT table_name
      FROM information_schema.tables
      WHERE table_schema='public'
      ORDER BY table_name
    `);
    res.json(result.rows.map(r => r.table_name));
  } catch (err) {
    console.error("DB Error:", err);
    res.status(500).json({ error: err.message });
  } finally { await client.end(); }
});

// --- Get paginated data safely ---
app.get("/data/:table", async (req, res) => {
  const { table } = req.params;
  const page = parseInt(req.query.page ) || 1;
  const limit = parseInt(req.query.limit ) || 5;
  const offset = (page - 1) * limit;

  const client = getClient();
  await client.connect();

  try {
    // Check if ID column exists
    const idCheck = await client.query(
      `SELECT column_name 
       FROM information_schema.columns 
       WHERE table_name = $1 AND column_name='ID'`,
      [table]
    );
    const hasID = idCheck.rows.length > 0;

    // Count total rows
    const totalResult = await client.query(`SELECT COUNT(*) FROM "${table}"`);
    const totalRows = parseInt(totalResult.rows[0].count);

    // Fetch data
    const orderBy = hasID ? `"ID"` : null;
    const result = await client.query(
      orderBy
        ? `SELECT * FROM "${table}" ORDER BY "ID"::int LIMIT $1 OFFSET $2`
        : `SELECT * FROM "${table}" LIMIT $1 OFFSET $2`,
      [limit, offset]
    );

    res.json({
      data: result.rows,
      page,
      limit,
      totalRows,
      totalPages: Math.ceil(totalRows / limit),
    });
  } catch (err) {
    console.error("DB Error:", err);
    res.status(500).json({ error: err.message });
  } finally { await client.end(); }
});

// --- Delete table ---
app.delete("/table/:name", async (req, res) => {
  const { name } = req.params;
  const client = getClient();
  await client.connect();
  try {
    await client.query(`DROP TABLE IF EXISTS "${name}"`);
    res.json({ message: `Table '${name}' deleted successfully` });
  } catch (err) {
    console.error("Delete Error:", err);
    res.status(500).json({ error: err.message });
  } finally { await client.end(); }
});

// --- Export table as Excel ---
app.get("/table/:table", async (req, res) => {
  const { table } = req.params;
  const client = getClient();
  await client.connect();

  try {
    const result = await client.query(`SELECT * FROM "${table}"`);
    if (result.rows.length === 0) return res.status(404).send("No data found");

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(table);

    worksheet.addRow(Object.keys(result.rows[0]));
    result.rows.forEach(row => worksheet.addRow(Object.values(row)));

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=${table}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Export Error:", err);
    res.status(500).send(err.message);
  } finally { await client.end(); }
});

// --- Update row ---
app.put("/data/:table/:id", async (req, res) => {
  const { table, id } = req.params;
  const updates = req.body;
  const client = getClient();
  await client.connect();

  try {
    if (!updates || Object.keys(updates).length === 0)
      return res.status(400).json({ error: "No fields provided for update" });

    const setClause = Object.keys(updates)
      .map((col, i) => `"${col}" = $${i + 1}`)
      .join(", ");
    const values = Object.values(updates);
    const sql = `UPDATE "${table}" SET ${setClause} WHERE "ID" = $${values.length + 1} RETURNING *`;

    const result = await client.query(sql, [...values, id]);

    if (result.rowCount === 0)
      return res.status(404).json({ error: `Row with ID=${id} not found` });

    res.json({ message: `Row updated`, row: result.rows[0] });
  } catch (err) {
    console.error("Update Error:", err);
    res.status(500).json({ error: err.message });
  } finally { await client.end(); }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
