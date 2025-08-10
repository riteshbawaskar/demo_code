import React from "react";

export default function JsonTable({ jsonString }) {
  let data = [];

  try {
    const parsed = JSON.parse(jsonString);
    if (Array.isArray(parsed)) {
      data = parsed;
    } else {
      data = [parsed]; // single object case
    }
  } catch (e) {
    return <div>Invalid JSON</div>;
  }

  if (!data.length) {
    return <div>No data</div>;
  }

  const columns = Array.from(
    new Set(data.flatMap((item) => Object.keys(item)))
  );

  return (
    <table border="1">
      <thead>
        <tr>
          {columns.map((col) => (
            <th key={col}>{col}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {data.map((row, idx) => (
          <tr key={idx}>
            {columns.map((col) => (
              <td key={col}>{row[col] ?? ""}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
