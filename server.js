const path = require("node:path");
const crypto = require("node:crypto");
const express = require("express");
const mysql = require("mysql2/promise");
const ExcelJS = require("exceljs");
const {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} = require("docx");

try {
  require("dotenv").config();
} catch (_) {
  // dotenv is optional in the managed runtime; platform variables are already loaded.
}

const app = express();
const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);
const SESSION_COOKIE = "wow_admin_session";
const SESSION_TTL_MS = 12 * 60 * 60 * 1000;
let pool;

app.disable("x-powered-by");
app.use(express.json({ limit: "64kb" }));
app.use(express.urlencoded({ extended: false, limit: "16kb" }));

function requiredText(value, field, maxLength = 500) {
  const text = String(value ?? "").trim();
  if (!text || text.length > maxLength) {
    const error = new Error(`${field} is required`);
    error.status = 400;
    throw error;
  }
  return text;
}

function optionalText(value, maxLength = 500) {
  return String(value ?? "").trim().slice(0, maxLength);
}

function parseDatabaseUrl() {
  if (!process.env.DATABASE_URL) {
    throw new Error("DATABASE_URL is not configured");
  }
  const url = new URL(process.env.DATABASE_URL);
  return {
    host: url.hostname,
    port: Number(url.port || 3306),
    user: decodeURIComponent(url.username),
    password: decodeURIComponent(url.password),
    database: decodeURIComponent(url.pathname.replace(/^\//, "")),
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    charset: "utf8mb4",
    ssl: { rejectUnauthorized: false },
  };
}

async function initDatabase() {
  pool = mysql.createPool(parseDatabaseUrl());
  await pool.query(`
    CREATE TABLE IF NOT EXISTS orders (
      id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
      product_name VARCHAR(255) NOT NULL,
      product_name_en VARCHAR(255) NOT NULL DEFAULT '',
      category VARCHAR(100) NOT NULL DEFAULT '',
      price DECIMAL(10,2) NOT NULL DEFAULT 0,
      size VARCHAR(30) NOT NULL DEFAULT '',
      color VARCHAR(100) NOT NULL DEFAULT '',
      quantity INT UNSIGNED NOT NULL DEFAULT 1,
      customer_name VARCHAR(255) NOT NULL,
      phone VARCHAR(60) NOT NULL,
      address TEXT NOT NULL,
      measurements VARCHAR(500) NOT NULL,
      product_image TEXT NOT NULL,
      created_at DATETIME(3) NOT NULL DEFAULT CURRENT_TIMESTAMP(3),
      PRIMARY KEY (id),
      INDEX idx_orders_created_at (created_at)
    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
  `);
}

function makeSessionToken() {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) return null;
  const issuedAt = Date.now().toString();
  const signature = crypto
    .createHmac("sha256", password)
    .update(issuedAt)
    .digest("base64url");
  return `${issuedAt}.${signature}`;
}

function isValidSessionToken(token) {
  const password = process.env.ADMIN_PASSWORD;
  if (!password || !token) return false;
  const [issuedAt, signature] = token.split(".");
  if (!issuedAt || !signature || !/^\d+$/.test(issuedAt)) return false;
  const age = Date.now() - Number(issuedAt);
  if (age < 0 || age > SESSION_TTL_MS) return false;
  const expected = crypto
    .createHmac("sha256", password)
    .update(issuedAt)
    .digest("base64url");
  if (signature.length !== expected.length) return false;
  return crypto.timingSafeEqual(
    Buffer.from(signature),
    Buffer.from(expected),
  );
}

function readCookies(req) {
  return Object.fromEntries(
    String(req.headers.cookie || "")
      .split(";")
      .map((part) => part.trim().split("="))
      .filter(([key, value]) => key && value)
      .map(([key, ...value]) => [key, decodeURIComponent(value.join("="))]),
  );
}

function isAdmin(req) {
  return isValidSessionToken(readCookies(req)[SESSION_COOKIE]);
}

function requireAdmin(req, res, next) {
  if (!isAdmin(req)) {
    return res.status(401).json({ ok: false, error: "admin_auth_required" });
  }
  next();
}

function adminCookieOptions(maxAge) {
  const hosted = Boolean(process.env.MANUS_PROJECT_ID) || process.env.FORCE_SECURE_COOKIES === "1";
  return `${SESSION_COOKIE}=VALUE; Max-Age=${maxAge}; Path=/; HttpOnly; SameSite=${hosted ? "None" : "Lax"}${hosted ? "; Secure" : ""}`;
}

function orderFromBody(body) {
  const price = Number(body.price);
  const quantity = Math.max(1, Math.min(99, Number(body.quantity) || 1));
  if (!Number.isFinite(price) || price < 0 || price > 100000000) {
    const error = new Error("price must be a valid number");
    error.status = 400;
    throw error;
  }
  return {
    product_name: requiredText(body.productName, "productName", 255),
    product_name_en: optionalText(body.productNameEn, 255),
    category: optionalText(body.category, 100),
    price,
    size: optionalText(body.size, 30),
    color: optionalText(body.color, 100),
    quantity,
    customer_name: requiredText(body.name, "name", 255),
    phone: requiredText(body.phone, "phone", 60),
    address: requiredText(body.address, "address", 2000),
    measurements: requiredText(body.measurements, "measurements", 500),
    product_image: optionalText(body.productImage, 2000),
  };
}

async function listOrders() {
  const [rows] = await pool.query(`
    SELECT id, product_name, product_name_en, category, price, size, color,
           quantity, customer_name, phone, address, measurements, product_image,
           created_at
    FROM orders
    ORDER BY id DESC
  `);
  return rows;
}

async function deleteOrders(ids) {
  if (!ids.length) return;
  await pool.query("DELETE FROM orders WHERE id IN (?)", [ids]);
}

function fileTimestamp() {
  return new Date().toISOString().replace(/[T:.Z]/g, "-").replace(/-+$/, "");
}

function setDownloadHeaders(res, filename, contentType) {
  res.setHeader("Content-Type", contentType);
  res.setHeader(
    "Content-Disposition",
    `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
  );
  res.setHeader("Cache-Control", "no-store");
}

async function buildExcel(orders) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WOW Fashion";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("الطلبات");
  sheet.columns = [
    { header: "رقم الطلب", key: "id", width: 12 },
    { header: "التاريخ", key: "created_at", width: 22 },
    { header: "المنتج", key: "product_name", width: 26 },
    { header: "الاسم بالإنجليزية", key: "product_name_en", width: 28 },
    { header: "التصنيف", key: "category", width: 14 },
    { header: "السعر", key: "price", width: 14 },
    { header: "المقاس", key: "size", width: 12 },
    { header: "اللون", key: "color", width: 16 },
    { header: "الكمية", key: "quantity", width: 10 },
    { header: "اسم العميل", key: "customer_name", width: 24 },
    { header: "الهاتف", key: "phone", width: 18 },
    { header: "العنوان", key: "address", width: 38 },
    { header: "المقاس بالتفصيل", key: "measurements", width: 34 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF15192B" } };
  sheet.views = [{ rightToLeft: true }];
  for (const order of orders) {
    sheet.addRow({
      ...order,
      created_at: new Date(order.created_at).toLocaleString("ar-EG"),
      price: Number(order.price),
    });
  }
  sheet.getColumn("price").numFmt = '#,##0.00 "ج.م"';
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) row.alignment = { vertical: "top", wrapText: true };
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function textCell(text, bold = false) {
  return new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text: String(text ?? ""), bold })] })],
  });
}

async function buildWord(orders) {
  const children = [
    new Paragraph({ text: "WOW Fashion — طلبات العملاء", heading: HeadingLevel.TITLE }),
    new Paragraph({ text: `عدد الطلبات: ${orders.length}` }),
    new Paragraph({ text: "" }),
  ];
  for (const order of orders) {
    children.push(
      new Paragraph({ text: `طلب رقم ${order.id}`, heading: HeadingLevel.HEADING_2 }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [textCell("البيان", true), textCell("التفاصيل", true)] }),
          new TableRow({ children: [textCell("التاريخ"), textCell(new Date(order.created_at).toLocaleString("ar-EG"))] }),
          new TableRow({ children: [textCell("المنتج"), textCell(`${order.product_name} — ${order.product_name_en}`)] }),
          new TableRow({ children: [textCell("السعر والكمية"), textCell(`${order.price} ج.م × ${order.quantity}`)] }),
          new TableRow({ children: [textCell("المقاس واللون"), textCell(`${order.size} — ${order.color}`)] }),
          new TableRow({ children: [textCell("العميل"), textCell(`${order.customer_name} — ${order.phone}`)] }),
          new TableRow({ children: [textCell("العنوان"), textCell(order.address)] }),
          new TableRow({ children: [textCell("المقاس بالتفصيل"), textCell(order.measurements)] }),
        ],
      }),
      new Paragraph({ text: "" }),
    );
  }
  return Packer.toBuffer(new Document({ sections: [{ properties: {}, children }] }));
}

async function exportOrders(res, format) {
  const orders = await listOrders();
  if (!orders.length) {
    return res.status(404).json({ ok: false, error: "no_orders" });
  }
  const ids = orders.map((order) => order.id);
  const stamp = fileTimestamp();
  const isExcel = format === "xlsx";
  const buffer = isExcel ? await buildExcel(orders) : await buildWord(orders);
  const filename = `wow-fashion-orders-${stamp}.${format}`;
  setDownloadHeaders(
    res,
    filename,
    isExcel
      ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  );
  res.once("finish", () => {
    deleteOrders(ids).catch((error) => console.error("Could not clear exported orders", error));
  });
  res.send(buffer);
}

app.get("/", (_req, res) => res.sendFile(path.join(ROOT, "index.html")));
app.get("/index.html", (_req, res) => res.sendFile(path.join(ROOT, "index.html")));
app.get("/admin", (_req, res) => res.sendFile(path.join(ROOT, "admin.html")));
app.get("/admin.html", (_req, res) => res.sendFile(path.join(ROOT, "admin.html")));
app.get("/manus-routes.json", (_req, res) => res.sendFile(path.join(ROOT, "manus-routes.json")));
app.get("/_app/health", (_req, res) => res.json({ ok: true, service: "wow-fashion" }));

app.post("/api/orders", async (req, res, next) => {
  try {
    const order = orderFromBody(req.body || {});
    const [result] = await pool.execute(
      `INSERT INTO orders
        (product_name, product_name_en, category, price, size, color, quantity,
         customer_name, phone, address, measurements, product_image)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        order.product_name,
        order.product_name_en,
        order.category,
        order.price,
        order.size,
        order.color,
        order.quantity,
        order.customer_name,
        order.phone,
        order.address,
        order.measurements,
        order.product_image,
      ],
    );
    res.status(201).json({ ok: true, orderId: result.insertId });
  } catch (error) {
    next(error);
  }
});

app.post("/api/admin/login", (req, res) => {
  if (!process.env.ADMIN_PASSWORD) {
    return res.status(503).json({ ok: false, error: "ADMIN_PASSWORD is not configured" });
  }
  const suppliedPassword = Buffer.from(String(req.body?.password || ""));
  const expectedPassword = Buffer.from(process.env.ADMIN_PASSWORD);
  const passwordMatches = suppliedPassword.length === expectedPassword.length && crypto.timingSafeEqual(suppliedPassword, expectedPassword);
  if (!passwordMatches) {
    return res.status(401).json({ ok: false, error: "invalid_password" });
  }
  const token = makeSessionToken();
  res.setHeader("Set-Cookie", adminCookieOptions(SESSION_TTL_MS / 1000).replace("VALUE", encodeURIComponent(token)));
  res.json({ ok: true });
});

app.post("/api/admin/logout", (req, res) => {
  res.setHeader("Set-Cookie", adminCookieOptions(0).replace("VALUE", ""));
  res.json({ ok: true });
});

app.get("/api/admin/status", (req, res) => {
  res.json({ ok: true, authenticated: isAdmin(req), configured: Boolean(process.env.ADMIN_PASSWORD) });
});

app.get("/api/admin/orders", requireAdmin, async (_req, res, next) => {
  try {
    res.json({ ok: true, orders: await listOrders() });
  } catch (error) {
    next(error);
  }
});

app.get("/api/admin/export/:format", requireAdmin, async (req, res, next) => {
  try {
    if (!["xlsx", "docx"].includes(req.params.format)) {
      return res.status(400).json({ ok: false, error: "unsupported_format" });
    }
    await exportOrders(res, req.params.format);
  } catch (error) {
    next(error);
  }
});

app.use((error, _req, res, _next) => {
  console.error(error);
  if (res.headersSent) return;
  res.status(error.status || 500).json({ ok: false, error: "server_error" });
});

async function start() {
  await initDatabase();
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`WOW Fashion server listening on ${PORT}`);
  });
}

start().catch((error) => {
  console.error("WOW Fashion server failed to start", error);
  process.exit(1);
});
