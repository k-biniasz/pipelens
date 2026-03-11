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
  if (req.method === "POST") {
    // Upsert an analysis result
    const { sessionId, analysisType, aeOwner, content } = req.body;
    if (!sessionId || !analysisType || !content) {
      return res.status(400).json({ error: "sessionId, analysisType, and content required" });
    }

    const conn = getConn();
    try {
      await conn.execute(
        `INSERT INTO analyses (session_id, analysis_type, ae_owner, content)
         VALUES (?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE content=VALUES(content), created_at=CURRENT_TIMESTAMP`,
        [sessionId, analysisType, aeOwner || null, content]
      );
      return res.status(200).json({ ok: true });
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  if (req.method === "GET") {
    // Fetch all analyses for a session
    const { sessionId } = req.query;
    if (!sessionId) return res.status(400).json({ error: "sessionId required" });

    const conn = getConn();
    try {
      const rows = await conn.execute(
        "SELECT analysis_type, ae_owner, content FROM analyses WHERE session_id = ?",
        [sessionId]
      );
      return res.status(200).json(rows);
    } catch (err) {
      return res.status(500).json({ error: err.message });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
}
