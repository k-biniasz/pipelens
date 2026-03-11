import { connect } from "@tidbcloud/serverless";

function getConn() {
  return connect({
    host:     process.env.TIDB_HOST,
    username: process.env.TIDB_USERNAME,
    password: process.env.TIDB_PASSWORD,
    database: process.env.TIDB_DATABASE,
  });
}

export default async function handler(req, res) {
  if (req.method === "GET") {
    // List 20 most recent sessions
    const conn = getConn();
    try {
      const rows = await conn.execute(
        "SELECT id, name, file_name, row_count, created_at FROM sessions ORDER BY created_at DESC LIMIT 20"
      );
      return res.status(200).json(rows);
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  if (req.method === "POST") {
    // Create session + bulk-insert rows
    const { id, name, fileName, colMap, companyContext, rows } = req.body;
    if (!id || !rows) return res.status(400).json({ error: "id and rows required" });

    const conn = getConn();
    try {
      // Upsert session metadata
      await conn.execute(
        `INSERT INTO sessions (id, name, file_name, col_map, company_context, row_count)
         VALUES (?, ?, ?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE
           name=VALUES(name), file_name=VALUES(file_name),
           col_map=VALUES(col_map), company_context=VALUES(company_context),
           row_count=VALUES(row_count), updated_at=CURRENT_TIMESTAMP`,
        [id, name || fileName, fileName, JSON.stringify(colMap), companyContext || "", rows.length]
      );

      // Delete old rows for this session (re-upload case)
      await conn.execute("DELETE FROM session_rows WHERE session_id = ?", [id]);

      // Bulk insert rows in chunks of 200
      const CHUNK = 200;
      for (let i = 0; i < rows.length; i += CHUNK) {
        const chunk = rows.slice(i, i + CHUNK);
        const placeholders = chunk.map(() => "(?, ?)").join(", ");
        const values = chunk.flatMap(r => [id, JSON.stringify(r)]);
        await conn.execute(
          `INSERT INTO session_rows (session_id, row_data) VALUES ${placeholders}`,
          values
        );
      }

      return res.status(200).json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  // GET session rows for a specific session
  if (req.method === "PUT") {
    const { id } = req.body;
    if (!id) return res.status(400).json({ error: "id required" });

    const conn = getConn();
    try {
      const [session] = await conn.execute(
        "SELECT id, name, file_name, col_map, company_context, row_count FROM sessions WHERE id = ?",
        [id]
      );
      if (!session) return res.status(404).json({ error: "Session not found" });

      const rowResults = await conn.execute(
        "SELECT row_data FROM session_rows WHERE session_id = ? ORDER BY id",
        [id]
      );
      const rows = rowResults.map(r =>
        typeof r.row_data === "string" ? JSON.parse(r.row_data) : r.row_data
      );

      return res.status(200).json({
        id: session.id,
        name: session.name,
        fileName: session.file_name,
        colMap: typeof session.col_map === "string" ? JSON.parse(session.col_map) : session.col_map,
        companyContext: session.company_context,
        rows,
      });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
}
