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
  origin: ["http://localhost:3000", "https://excel-front-navy.vercel.app"],
  methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// --- Upload Setup ---
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

// --- Home ---
app.get("/", (req, res) => res.send("Backend is running!"));

// --- Upload Excel ---
app.post("/upload", upload.single("excel"), async (req, res) => {
  if (!req.file) return res.status(400).send("No file uploaded");

  let { tableName } = req.body;
  const rows = await readXlsxFile(req.file.path);
  if (rows.length < 2) return res.status(400).send("Excel has no data");

  const headers = rows[0]; // Keep all columns including Excel's ID
  const dataRows = rows.slice(1);

  if (!tableName) tableName = path.parse(req.file.originalname).name;

  const client = getClient();
  await client.connect();

  try {
    // Check existing columns
    const tableCheck = await client.query(
      `SELECT column_name FROM information_schema.columns WHERE table_name = $1`,
      [tableName]
    );
    const existingColumns = tableCheck.rows.map(r => r.column_name);

    // Ensure consistent lowercase id column: table_id
    const idColumnName = "table_id";

    if (tableCheck.rows.length === 0) {
      // Table does not exist — create with table_id
      const columnsSQL = headers.map(h => `"${h}" TEXT`).join(", ");
      const createSQL = `
        CREATE TABLE "${tableName}" (
          "${idColumnName}" SERIAL PRIMARY KEY,
          ${columnsSQL}
        )
      `;
      await client.query(createSQL);
    } else if (!existingColumns.includes(idColumnName)) {
      // Table exists but no table_id — add it
      await client.query(`ALTER TABLE "${tableName}" ADD COLUMN "${idColumnName}" SERIAL PRIMARY KEY`);
    }

    // Insert rows (skip table_id)
    for (let row of dataRows) {
      const placeholders = headers.map((_, idx) => `$${idx + 1}`).join(", ");
      const insertSQL = `INSERT INTO "${tableName}" (${headers.map(h => `"${h}"`).join(", ")}) VALUES (${placeholders})`;
      await client.query(insertSQL, row);
    }

    res.json({ message: "Excel inserted successfully", table: tableName, inserted: dataRows.length });
  } catch (err) {
    console.error("Upload Error:", err);
    res.status(500).send("Error: " + err.message);
  } finally {
    await client.end();
  }
});

// --- List Tables ---
app.get("/tables", async (req, res) => {
  const client = getClient();
  await client.connect();
  try {
    const result = await client.query(`SELECT table_name FROM information_schema.tables 
      WHERE table_schema='public'`);
    res.json(result.rows.map(r => r.table_name));
  } catch (err) {
    console.error("DB Error:", err);
    res.status(500).json({ error: err.message });
  } finally { await client.end(); }
});

// --- Get Paginated Data ---
app.get("/data/:table", async (req, res) => {
  const { table } = req.params;
  const page = parseInt(req.query.page ) || 1;
  const limit = parseInt(req.query.limit ) || 5;
  const offset = (page - 1) * limit;

  const client = getClient();
  await client.connect();

  try {
    // Detect ID column dynamically
    const colResult = await client.query(
      `SELECT column_name FROM information_schema.columns WHERE table_name = $1 AND column_name ILIKE '%id%' ORDER BY ordinal_position LIMIT 1`,
      [table]
    );
    if (colResult.rows.length === 0)
      return res.status(400).json({ error: "No ID column found in table" });
    const idColumn = colResult.rows[0].column_name;

    const totalResult = await client.query(`SELECT COUNT(*) FROM "${table}"`);
    const totalRows = parseInt(totalResult.rows[0].count);

    const result = await client.query(
      `SELECT * FROM "${table}" ORDER BY "${idColumn}" LIMIT $1 OFFSET $2`,
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

// --- Delete Table ---
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

// --- Export Table ---
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

// --- Update Row ---
app.put("/data/:table/:id", async (req, res) => {
  const { table, id } = req.params;
  const updates = req.body;
  const client = getClient();
  await client.connect();

  try {
    // Validate ID
    if (!id || isNaN(Number(id))) {
      return res.status(400).json({ error: "Invalid or missing row ID" });
    }

    if (!updates || Object.keys(updates).length === 0) {
      return res.status(400).json({ error: "No fields provided for update" });
    }

    // Build SET clause
    const setClause = Object.keys(updates)
      .map((col, i) => `"${col}" = $${i + 1}`)
      .join(", ");
    const values = Object.values(updates);

    // Add id at the end
    values.push(Number(id));

    const sql = `UPDATE "${table}" SET ${setClause} WHERE table_id = $${values.length} RETURNING *;`;
    const result = await client.query(sql, values);

    if (result.rowCount === 0) {
      return res.status(404).json({ error: `Row with table_id=${id} not found` });
    }

    res.json({ message: "Row updated", row: result.rows[0] });
  } catch (err) {
    console.error("Update Error:", err);
    res.status(500).json({ error: err.message });
  } finally {
    await client.end();
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
