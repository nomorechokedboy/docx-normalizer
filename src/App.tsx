import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";
import { renderAsync } from "docx-preview";

type Margins = {
  top: number;
  right: number;
  bottom: number;
  left: number;
};

type SectionInfo = {
  index: number;
  margins: Margins | null;
};

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const presets: Record<string, Margins> = {
  "vn-admin": { top: 25, right: 20, bottom: 25, left: 30 },
  balanced: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
  compact: { top: 15, right: 15, bottom: 15, left: 15 },
};

function twips(mm: number) {
  return String(Math.round((mm / 25.4) * 1440));
}

function mmFromTwips(value: string | null) {
  if (!value || !/^[-+]?\d+$/.test(value)) return null;

  return Math.round((Number(value) / 1440) * 25.4 * 10) / 10;
}

function getSectPrs(xml: Document) {
  return Array.from(xml.getElementsByTagNameNS(WORD_NS, "sectPr"));
}

function readSections(xmlText: string): SectionInfo[] {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");

  if (xml.getElementsByTagName("parsererror").length) {
    throw new Error("The document XML could not be parsed.");
  }

  return getSectPrs(xml).map((section, i) => {
    const pgMar = Array.from(
      section.getElementsByTagNameNS(WORD_NS, "pgMar"),
    )[0];

    if (!pgMar) {
      return { index: i + 1, margins: null };
    }

    const read = (name: string) =>
      mmFromTwips(pgMar.getAttributeNS(WORD_NS, name));

    const values = {
      top: read("top"),
      right: read("right"),
      bottom: read("bottom"),
      left: read("left"),
    };

    if (Object.values(values).some((value) => value === null)) {
      return { index: i + 1, margins: null };
    }

    return {
      index: i + 1,
      margins: values as Margins,
    };
  });
}

async function reconstructWithMargins(file: File, margins: Margins) {
  const zip = await JSZip.loadAsync(file);
  const docEntry = zip.file("word/document.xml");

  if (!docEntry) {
    throw new Error(
      "This file does not contain word/document.xml and may not be a valid DOCX.",
    );
  }

  const sourceXml = await docEntry.async("string");
  const xml = new DOMParser().parseFromString(sourceXml, "application/xml");

  if (xml.getElementsByTagName("parsererror").length) {
    throw new Error("The DOCX main document XML is malformed.");
  }

  const sections = getSectPrs(xml);

  if (!sections.length) {
    throw new Error("No section properties were found in this document.");
  }

  for (const section of sections) {
    let pgMar = Array.from(section.children).find(
      (el) => el.namespaceURI === WORD_NS && el.localName === "pgMar",
    );

    if (!pgMar) {
      pgMar = xml.createElementNS(WORD_NS, "w:pgMar");

      const pageSize = Array.from(section.children).find(
        (el) => el.namespaceURI === WORD_NS && el.localName === "pgSz",
      );

      if (pageSize?.nextSibling) {
        section.insertBefore(pgMar, pageSize.nextSibling);
      } else if (pageSize) {
        section.appendChild(pgMar);
      } else {
        section.insertBefore(pgMar, section.firstChild);
      }
    }

    for (const key of ["top", "right", "bottom", "left"] as const) {
      pgMar.setAttributeNS(WORD_NS, `w:${key}`, twips(margins[key]));
    }
  }

  const serialized = new XMLSerializer().serializeToString(xml);

  // Rebuild a new ZIP package from all original entries.
  // Only document.xml is changed; other OOXML parts are retained.
  zip.file("word/document.xml", serialized);

  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return {
    blob,
    sections: sections.length,
    originalXml: sourceXml,
    updatedXml: serialized,
  };
}

function Preview({ blob, label }: { blob: Blob | null; label: string }) {
  const host = useRef<HTMLDivElement>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    let active = true;

    if (!host.current || !blob) return;

    host.current.innerHTML = "";
    setError("");

    blob
      .arrayBuffer()
      .then((buffer) =>
        renderAsync(buffer, host.current!, undefined, {
          className: "docx-preview",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
        }),
      )
      .catch((err) => {
        if (active) {
          setError(
            err instanceof Error
              ? err.message
              : "Preview could not be rendered.",
          );
        }
      });

    return () => {
      active = false;

      if (host.current) {
        host.current.innerHTML = "";
      }
    };
  }, [blob]);

  return (
    <section className="min-w-0 overflow-hidden rounded-[10px] border border-[#e7eaf0] bg-[#f3f4f7]">
      <div className="flex h-[42px] items-center justify-between gap-[6px] border-b border-[#e7eaf0] bg-white px-[11px] text-[10px] font-bold">
        <span>{label}</span>

        <span className="text-[9px] font-normal text-[#9aa1ad]">
          Word layout preview
        </span>
      </div>

      {error && (
        <div className="p-[10px] text-[10px] text-[#a13d3d]">
          Preview error: {error}
        </div>
      )}

      {!blob && (
        <div className="flex min-h-[300px] flex-col items-center justify-center gap-2 px-[25px] py-[25px] text-center text-[#8d95a2]">
          <div className="text-[34px] text-[#bdc5d7]">▤</div>

          <strong className="text-[11px] text-[#596273]">
            Preview will appear here
          </strong>

          <span className="text-[10px]">
            Choose a DOCX file to get started.
          </span>
        </div>
      )}

      <div
        ref={host}
        className="
          max-h-[900px] min-h-[300px] overflow-auto p-[13px]
          [&:empty]:hidden
          [&_.docx-wrapper]:!bg-transparent
          [&_.docx-wrapper]:!p-0
          [&_.docx-wrapper>section]:!mx-auto
          [&_.docx-wrapper>section]:!mb-[12px]
          [&_.docx-wrapper>section]:!max-w-full
          [&_.docx-wrapper>section]:!shadow-[0_2px_8px_#1c29400f]
          [&_.docx]:!max-w-full
          max-[1050px]:max-h-[600px]
          max-[720px]:max-h-[650px]
        "
      />
    </section>
  );
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [originalBlob, setOriginalBlob] = useState<Blob | null>(null);
  const [outputBlob, setOutputBlob] = useState<Blob | null>(null);

  const [preset, setPreset] = useState("vn-admin");
  const [margins, setMargins] = useState<Margins>(presets["vn-admin"]);

  const [sections, setSections] = useState<SectionInfo[]>([]);
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);

  const [stats, setStats] = useState<{
    sections: number;
    preservedParts: number;
  } | null>(null);

  function choosePreset(value: string) {
    setPreset(value);

    if (value in presets) {
      setMargins(presets[value]);
    }

    setOutputBlob(null);
  }

  async function loadFile(next: File | undefined) {
    if (!next) return;

    if (!next.name.toLowerCase().endsWith(".docx")) {
      setStatus(
        "Please choose a .docx file. Legacy .doc files are not supported in this demo.",
      );
      return;
    }

    setFile(next);
    setOriginalBlob(next);
    setOutputBlob(null);
    setStats(null);
    setStatus("Inspecting document…");

    try {
      const zip = await JSZip.loadAsync(next);
      const entry = zip.file("word/document.xml");

      if (!entry) {
        throw new Error(
          "Could not find word/document.xml. Check that this is a valid DOCX file.",
        );
      }

      const xml = await entry.async("string");
      const found = readSections(xml);

      setSections(found);

      setStatus(
        `Loaded ${next.name} · ${found.length} section${found.length === 1 ? "" : "s"} detected.`,
      );
    } catch (err) {
      setStatus(
        err instanceof Error ? err.message : "Unable to inspect this file.",
      );

      setSections([]);
    }
  }

  async function normalize() {
    if (!file) return;

    setBusy(true);
    setStatus("Rebuilding package and preparing preview…");

    try {
      const result = await reconstructWithMargins(file, margins);

      // Reopen generated archive as a basic integrity check.
      const check = await JSZip.loadAsync(result.blob);

      if (
        !check.file("word/document.xml") ||
        !check.file("[Content_Types].xml")
      ) {
        throw new Error("Output package failed integrity checks.");
      }

      const count = Object.keys(check.files).filter(
        (name) => !check.files[name].dir,
      ).length;

      setOutputBlob(result.blob);

      setStats({
        sections: result.sections,
        preservedParts: count,
      });

      setStatus(
        "Candidate created. Review both previews before using the output; visual safety checks are not automated in this prototype.",
      );
    } catch (err) {
      setOutputBlob(null);

      setStatus(err instanceof Error ? err.message : "Normalization failed.");
    } finally {
      setBusy(false);
    }
  }

  function download() {
    if (!outputBlob || !file) return;

    const url = URL.createObjectURL(outputBlob);
    const a = document.createElement("a");

    a.href = url;
    a.download = `${file.name.replace(/\.docx$/i, "")}-normalized.docx`;
    a.click();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  const changed = outputBlob !== null;

  return (
    <main className="min-h-screen bg-[#f5f6f8] text-[#202735] [font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe_UI',sans-serif] [font-synthesis:none] [text-rendering:optimizeLegibility]">
      {/* Header */}
      <header className="flex h-[68px] items-center gap-3 border-b border-[#e7e9ee] bg-white px-[clamp(18px,4vw,60px)]">
        <div className="grid size-[34px] shrink-0 place-items-center rounded-[10px] bg-[#233fce] text-[19px] font-extrabold text-white shadow-[0_4px_10px_#233fce2b]">
          M
        </div>

        <div>
          <div className="text-[15px] font-extrabold tracking-[-0.3px]">
            Margin Lab
          </div>

          <div className="mt-[2px] text-[11px] text-[#8991a0]">
            DOCX normalization prototype
          </div>
        </div>

        <div className="ml-auto flex items-center gap-[7px] rounded-[30px] border border-[#dfeee6] bg-[#f4fbf6] px-[11px] py-[7px] text-[11px] text-[#34744c] max-[720px]:p-2 max-[720px]:text-[0px]">
          <span className="size-[6px] rounded-full bg-[#43a66b] max-[720px]:size-[7px]" />
          Runs locally in your browser
        </div>
      </header>

      <div className="mx-auto max-w-[1500px] px-[clamp(18px,4vw,60px)] pb-6 pt-9 max-[720px]:pt-[25px]">
        {/* Introduction */}
        <div className="mb-[27px] max-w-[680px]">
          <div className="mb-[10px] text-[10px] font-bold tracking-[1.7px] text-[#455fe1]">
            PROTOTYPE · 0.1
          </div>

          <h1 className="m-0 text-[clamp(29px,3vw,42px)] leading-[1.13] font-extrabold tracking-[-1.8px] text-[#192235] max-[720px]:tracking-[-1px]">
            Standardize margins.
            <br />
            <span className="text-[#5369dc]">Keep the document intact.</span>
          </h1>

          <p className="mt-[14px] mb-0 max-w-[640px] text-[13px] leading-[1.8] text-[#6f7786]">
            Rebuild a DOCX package while retaining its original content XML,
            media, styles and relationships. Change page margins, then compare
            the result side by side.
          </p>
        </div>

        {/* Workspace */}
        <div className="grid grid-cols-[310px_minmax(0,1fr)] items-start gap-5 max-[1050px]:grid-cols-[270px_minmax(0,1fr)] max-[1050px]:gap-[13px] max-[720px]:grid-cols-1">
          {/* Control panel */}
          <aside className="rounded-[15px] border border-[#e5e8ee] bg-white p-[21px] shadow-[0_5px_20px_#18233b06] max-[1050px]:p-4 max-[720px]:order-1">
            {/* Source document */}
            <div className="flex items-start gap-[11px]">
              <span className="grid size-[27px] shrink-0 place-items-center rounded-lg bg-[#eef1ff] text-[11px] font-bold text-[#4359d3]">
                01
              </span>

              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  Source document
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  Choose a Word file to inspect.
                </p>
              </div>
            </div>

            {/* Upload */}
            <label
              htmlFor="docx-file"
              className="mt-[18px] flex min-h-[142px] cursor-pointer flex-col items-center justify-center gap-[7px] rounded-[11px] border-[1.5px] border-dashed border-[#cdd3e3] bg-[#fafbff] p-[14px] text-center transition-colors duration-200 hover:border-[#6275e3] hover:bg-[#f6f7ff]"
            >
              <span className="grid size-[34px] place-items-center rounded-[10px] bg-[#e9edff] text-[23px] text-[#465bd2]">
                ↑
              </span>

              <strong className="max-w-full overflow-wrap-anywhere text-[12px]">
                {file ? file.name : "Drop a DOCX here"}
              </strong>

              <span className="text-[10px] text-[#8b93a1]">
                {file
                  ? `${(file.size / 1024).toFixed(0)} KB · selected`
                  : "or click to browse · .docx only"}
              </span>

              <input
                id="docx-file"
                type="file"
                accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                className="hidden"
                onChange={(e) => void loadFile(e.target.files?.[0])}
              />
            </label>

            {/* Section summary */}
            {sections.length > 0 && (
              <div className="mt-3 flex items-center gap-[10px] rounded-[9px] bg-[#f6f8ff] p-[11px]">
                <div className="font-[Manrope] text-[20px] font-extrabold text-[#435bd3]">
                  {sections.length}
                </div>

                <div>
                  <strong className="block text-[11px]">
                    {sections.length === 1 ? "Section found" : "Sections found"}
                  </strong>

                  <span className="mt-[3px] block text-[10px] text-[#7e8796]">
                    Margins are applied to every section.
                  </span>
                </div>
              </div>
            )}

            <div className="my-[22px] h-px bg-[#edf0f4]" />

            {/* Margin profile */}
            <div className="flex items-start gap-[11px]">
              <span className="grid size-[27px] shrink-0 place-items-center rounded-lg bg-[#eef1ff] text-[11px] font-bold text-[#4359d3]">
                02
              </span>

              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  Margin profile
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  All values are in millimeters.
                </p>
              </div>
            </div>

            <label
              htmlFor="profile"
              className="mt-[19px] mb-[7px] block text-[10px] font-bold text-[#737c8b]"
            >
              Preset
            </label>

            <select
              id="profile"
              value={preset}
              onChange={(e) => choosePreset(e.target.value)}
              className="w-full rounded-lg border border-[#dfe3eb] bg-white p-[10px] text-[11px] text-[#333c4b] outline-none focus:border-[#687be3] focus:shadow-[0_0_0_3px_#5369dc16]"
            >
              <option value="vn-admin">
                Administrative example · 25 / 20 / 25 / 30
              </option>
              <option value="balanced">Balanced · 25.4 on all sides</option>
              <option value="compact">Compact · 15 on all sides</option>
              <option value="custom">Custom values</option>
            </select>

            {/* Margin inputs */}
            <div className="mt-[13px] grid grid-cols-2 gap-[10px]">
              {(["top", "right", "bottom", "left"] as const).map((key) => (
                <label key={key}>
                  <span className="mb-[6px] block text-[10px] text-[#737c8b]">
                    {key[0].toUpperCase() + key.slice(1)}
                  </span>

                  <div className="flex items-center rounded-lg border border-[#e0e4ec] px-2">
                    <input
                      type="number"
                      min="0"
                      max="100"
                      step="0.5"
                      value={margins[key]}
                      onChange={(e) => {
                        setPreset("custom");

                        setMargins((old) => ({
                          ...old,
                          [key]: Math.max(
                            0,
                            Math.min(100, Number(e.target.value)),
                          ),
                        }));

                        setOutputBlob(null);
                      }}
                      className="w-full min-w-0 border-0 bg-transparent py-[9px] text-[12px] font-semibold text-[#283143] outline-none focus:shadow-none"
                    />

                    <span className="text-[10px] text-[#9ba2ae]">mm</span>
                  </div>
                </label>
              ))}
            </div>

            {/* Preservation note */}
            <div className="mt-[19px] flex gap-[9px] rounded-[9px] border border-[#e1f1e7] bg-[#f3faf6] px-[10px] py-3">
              <span className="grid size-5 shrink-0 place-items-center rounded-full bg-[#d8f1e1] text-[12px] text-[#27814b]">
                ✓
              </span>

              <div>
                <strong className="text-[10px] text-[#2d6743]">
                  Preservation policy
                </strong>

                <p className="mt-1 mb-0 text-[10px] leading-[1.6] text-[#62806c]">
                  Keep original text, tables, images, styles and relationships.
                  Only page margins are changed in this demo.
                </p>
              </div>
            </div>

            {/* Normalize button */}
            <button
              onClick={() => void normalize()}
              disabled={!file || busy || sections.length === 0}
              className="mt-[14px] flex w-full items-center justify-between rounded-[9px] border-0 bg-[#344fd0] px-[13px] py-3 text-[11px] font-semibold text-white shadow-[0_4px_10px_#344fd02b] transition-colors hover:bg-[#283fb9] disabled:cursor-not-allowed disabled:opacity-[0.45]"
            >
              {busy ? (
                <>
                  <span className="flex items-center gap-2">
                    <span className="size-[13px] animate-spin rounded-full border-2 border-white/40 border-t-white" />
                    Processing…
                  </span>
                </>
              ) : (
                <>
                  <span>Normalize margins</span>
                  <span>→</span>
                </>
              )}
            </button>

            {/* Download button */}
            {changed && (
              <button
                onClick={download}
                className="mt-[14px] flex w-full items-center justify-center gap-[9px] rounded-[9px] border border-[#d7deeb] bg-white px-[13px] py-3 text-[11px] font-semibold text-[#3449b5] hover:bg-[#f7f8ff]"
              >
                <span>↓</span>
                Download normalized DOCX
              </button>
            )}

            {/* Status */}
            {status && (
              <div
                role="status"
                className="mt-[11px] overflow-wrap-anywhere rounded-[7px] bg-[#f7f8fa] px-[10px] py-[9px] text-[10px] leading-[1.6] text-[#747e8f]"
              >
                {status}
              </div>
            )}

            {/* Integrity check */}
            {stats && (
              <div className="mt-[11px] rounded-[9px] border border-[#dcefe3] bg-[#f8fcf9] p-[11px]">
                <div className="flex items-center gap-[7px] text-[10px] text-[#2b7548]">
                  <span className="grid size-[17px] place-items-center rounded-full bg-[#d9f1e2]">
                    ✓
                  </span>

                  <strong>Package integrity check passed</strong>
                </div>

                <div className="mt-3 grid grid-cols-2 gap-2">
                  <div className="rounded-[7px] border border-[#e6f1e9] bg-white p-2">
                    <strong className="block font-[Manrope] text-[18px] font-bold text-[#2d6844]">
                      {stats.sections}
                    </strong>

                    <span className="mt-[3px] block text-[9px] text-[#7e8d83]">
                      Sections updated
                    </span>
                  </div>

                  <div className="rounded-[7px] border border-[#e6f1e9] bg-white p-2">
                    <strong className="block font-[Manrope] text-[18px] font-bold text-[#2d6844]">
                      {stats.preservedParts}
                    </strong>

                    <span className="mt-[3px] block text-[9px] text-[#7e8d83]">
                      Package entries retained
                    </span>
                  </div>
                </div>

                <p className="mt-[9px] mb-0 text-[9px] leading-[1.6] text-[#73877a]">
                  This verifies basic package integrity, not visual correctness.
                  Inspect the previews for table or image issues.
                </p>
              </div>
            )}

            {/* Privacy footer */}
            <div className="mt-[18px] flex items-center justify-center gap-[6px] text-[9px] text-[#959dab]">
              <span className="text-[16px]">⌁</span>
              Your file stays in this browser tab. Nothing is uploaded.
            </div>
          </aside>

          {/* Preview area */}
          <section className="min-w-0 rounded-[15px] border border-[#e5e8ee] bg-white p-5 shadow-[0_5px_20px_#18233b06] max-[1050px]:p-[14px] max-[720px]:order-2">
            <div className="mb-[17px] flex items-center justify-between gap-[10px]">
              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  Visual comparison
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  Original versus normalized candidate
                </p>
              </div>

              <div className="flex items-center gap-[6px] whitespace-nowrap text-[10px] text-[#737c8b]">
                <span className="ml-2 inline-block size-[7px] rounded-full bg-[#a6adb9]" />
                Original
                <span className="ml-[7px] inline-block size-[7px] rounded-full bg-[#4e65dc]" />
                Candidate
              </div>
            </div>

            {/* Preview cards */}
            <div className="grid grid-cols-2 items-start gap-3 max-[1050px]:grid-cols-1 max-[720px]:grid-cols-1">
              <Preview blob={originalBlob} label="Original" />

              <Preview blob={outputBlob} label="Normalized candidate" />
            </div>

            {/* Disclaimer */}
            <div className="mt-[14px] flex gap-[9px] rounded-[9px] border border-[#f3e7bf] bg-[#fff9e9] p-3 text-[#92712b]">
              <span className="text-[15px]">ⓘ</span>

              <p className="m-0 text-[10px] leading-[1.7] text-[#8c7845]">
                <strong className="text-[#765d22]">
                  Prototype limitation:
                </strong>{" "}
                browser previews are for visual inspection only. This demo does
                not automatically detect clipping, compare table geometry, or
                guarantee that Microsoft Word and LibreOffice will render
                identically.
              </p>
            </div>
          </section>
        </div>

        {/* Footer */}
        <footer className="mt-5 flex justify-between gap-[10px] px-[3px] text-[9px] text-[#a0a6b1] max-[720px]:flex-col">
          <span>Margin Lab · local prototype</span>
          <span>DOCX package reconstruction · not production-ready</span>
        </footer>
      </div>
    </main>
  );
}
