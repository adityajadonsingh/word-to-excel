"use client";

import { useState } from "react";

type UploadFile = {
  file: File;
  progress: number;
  status: "pending" | "uploading" | "done" | "error";
};

export default function Home() {
  const [files, setFiles] = useState<UploadFile[]>([]);
  const [uploading, setUploading] = useState(false);
  const MAX_FILES = 3;

  // ===== Handle file selection =====
  const onSelectFiles = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(e.target.files || []);

    if (selected.length + files.length > MAX_FILES) {
      alert("You can upload a maximum of 3 Word files at a time.");
      return;
    }

    const mapped = selected.map((f) => ({
      file: f,
      progress: 0,
      status: "pending" as const,
    }));

    setFiles((prev) => [...prev, ...mapped]);
    e.target.value = "";
  };

  // ===== Upload files one by one =====
  const uploadFiles = async () => {
    setUploading(true);

    const updated = [...files];

    for (let i = 0; i < updated.length; i++) {
      if (updated[i].status === "done") continue;

      updated[i].status = "uploading";
      setFiles([...updated]);

      try {
        const form = new FormData();
        form.append("file", updated[i].file);

        // Fake progress (realistic UX)
        let progress = 0;
        const interval = setInterval(() => {
          progress = Math.min(progress + 10, 90);
          updated[i].progress = progress;
          setFiles([...updated]);
        }, 300);

        const res = await fetch("http://localhost:8000/upload-docx", {
          method: "POST",
          body: form,
        });

        clearInterval(interval);

        if (!res.ok) throw new Error();

        updated[i].progress = 100;
        updated[i].status = "done";
        setFiles([...updated]);
      } catch {
        updated[i].status = "error";
        setFiles([...updated]);
      }
    }

    setUploading(false);
  };

  // ===== Download Excel =====
  const downloadExcel = async () => {
    const res = await fetch("http://localhost:8000/download-excel"); // GET

    if (!res.ok) {
      alert("No data available to download");
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "final_output.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
  };

  // ===== Reset =====
const reset = async () => {
  await fetch("http://localhost:8000/reset-session", {
    method: "POST"
  });
  setFiles([]);
};


  return (
    <div className="min-h-screen bg-gray-100 flex items-center justify-center px-4">
      <div className="w-full max-w-xl bg-white rounded-2xl shadow-lg p-6">
        <h1 className="text-2xl text-center font-bold text-gray-800 mb-4">
          Word → Excel Converter
        </h1>

        {/* File Input */}
        <label className="block mb-4">
          <input
            type="file"
            accept=".docx"
            multiple
            disabled={files.length >= MAX_FILES || uploading}
            onChange={onSelectFiles}
            className="hidden"
          />
          <div className="cursor-pointer border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-blue-500 transition">
            <p className="text-gray-600">Click to select Word files (max 3)</p>
          </div>
        </label>

        {/* File List */}
        <div className="space-y-3 mb-4">
          {files.map((f, i) => (
            <div key={i} className="border rounded-lg p-3">
              <div className="flex justify-between text-sm mb-1">
                <span className="truncate">{f.file.name}</span>
                <span className="text-gray-500">
                  {f.status === "done"
                    ? "Done"
                    : f.status === "uploading"
                      ? "Uploading..."
                      : f.status === "error"
                        ? "Error"
                        : "Pending"}
                </span>
              </div>

              <div className="w-full bg-gray-200 rounded-full h-2">
                <div
                  className={`h-2 rounded-full transition-all ${
                    f.status === "error" ? "bg-red-500" : "bg-blue-500"
                  }`}
                  style={{ width: `${f.progress}%` }}
                />
              </div>
            </div>
          ))}
        </div>

        {/* Actions */}
        <div className="flex gap-3">
          <button
            onClick={uploadFiles}
            disabled={uploading || files.length === 0}
            className="flex-1 cursor-pointer bg-blue-600 text-white py-2 rounded-lg hover:bg-blue-700 disabled:opacity-50"
          >
            {uploading ? "Processing..." : "Start Upload"}
          </button>

          <button
            onClick={reset}
            disabled={uploading}
            className="flex-1 cursor-pointer border text-gray-700 border-gray-300 py-2 rounded-lg hover:bg-gray-100"
          >
            Clear
          </button>
        </div>

        {/* After Upload */}
        {files.length > 0 && files.every((f) => f.status === "done") && (
          <div className="mt-6 space-y-3">
            <button
              onClick={downloadExcel}
              className="w-full cursor-pointer bg-green-600 text-white py-2 rounded-lg hover:bg-green-700"
            >
              Download Excel
            </button>

            <button
              onClick={reset}
              className="w-full cursor-pointer border py-2 rounded-lg hover:bg-gray-100"
            >
              Upload More Files
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
