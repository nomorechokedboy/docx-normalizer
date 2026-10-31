import { DocxEditor } from "@docx-editor.dev/react";
import "@docx-editor.dev/core/styles/editor.css";
import JSZip from "jszip";
import { useEffect, useState } from "react";

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
    throw new Error("Không thể phân tích cấu trúc XML của tài liệu.");
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
      "Không tìm thấy word/document.xml. Tệp có thể không phải là tài liệu DOCX hợp lệ.",
    );
  }

  const sourceXml = await docEntry.async("string");
  const xml = new DOMParser().parseFromString(sourceXml, "application/xml");

  if (xml.getElementsByTagName("parsererror").length) {
    throw new Error("Cấu trúc XML chính của tài liệu DOCX không hợp lệ.");
  }

  const sections = getSectPrs(xml);

  if (!sections.length) {
    throw new Error("Không tìm thấy thông tin phân đoạn trong tài liệu.");
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

  // Tạo lại gói ZIP từ toàn bộ các thành phần ban đầu.
  // Chỉ thay đổi document.xml, các thành phần OOXML khác được giữ nguyên.
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

function Preview({
  blob,
  label,
  title,
}: {
  blob: Blob | null;
  label: string;
  title?: string;
}) {
  const [bytes, setBytes] = useState<Uint8Array | null>(null);
  const [error, setError] = useState("");

  // Đọc Blob thành byte để truyền cho docx-editor. Mỗi lần blob đổi,
  // một Uint8Array mới được tạo nên trình xem sẽ tự nạp lại tài liệu.
  useEffect(() => {
    let active = true;

    setError("");

    if (!blob) {
      setBytes(null);
      return;
    }

    blob
      .arrayBuffer()
      .then((buffer) => {
        if (active) setBytes(new Uint8Array(buffer));
      })
      .catch((err) => {
        if (active) {
          setBytes(null);
          setError(
            err instanceof Error
              ? err.message
              : "Không thể hiển thị bản xem trước.",
          );
        }
      });

    return () => {
      active = false;
    };
  }, [blob]);

  return (
    <section className="min-w-0 overflow-hidden rounded-[10px] border border-[#e7eaf0] bg-[#f3f4f7]">
      <div className="flex h-[42px] items-center justify-between gap-[6px] border-b border-[#e7eaf0] bg-white px-[11px] text-[10px] font-bold">
        <span>{label}</span>

        <span className="text-[9px] font-normal text-[#9aa1ad]">
          Bản xem trước bố cục Word
        </span>
      </div>

      {error && (
        <div className="p-[10px] text-[10px] text-[#a13d3d]">
          Lỗi hiển thị: {error}
        </div>
      )}

      {!blob && (
        <div className="flex min-h-[300px] flex-col items-center justify-center gap-2 px-[25px] py-[25px] text-center text-[#8d95a2]">
          <div className="text-[34px] text-[#bdc5d7]">▤</div>

          <strong className="text-[11px] text-[#596273]">
            Bản xem trước sẽ hiển thị tại đây
          </strong>

          <span className="text-[10px]">Chọn một tệp DOCX để bắt đầu.</span>
        </div>
      )}

      {bytes && (
        // docx-editor cần một vùng chứa có chiều cao xác định để cuộn trang.
        <div className="h-[600px] min-h-[300px] overflow-hidden max-[1050px]:h-[560px] max-[720px]:h-[500px]">
          <DocxEditor
            document={bytes}
            mode="view"
            openFailure={false}
            menu={false}
            rulers={false}
            navigation={false}
            title={title}
          />
        </div>
      )}
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
        "Vui lòng chọn tệp .docx. Định dạng .doc cũ hiện chưa được hỗ trợ.",
      );
      return;
    }

    setFile(next);
    setOriginalBlob(next);
    setOutputBlob(null);
    setStats(null);
    setStatus("Đang kiểm tra tài liệu…");

    try {
      const zip = await JSZip.loadAsync(next);
      const entry = zip.file("word/document.xml");

      if (!entry) {
        throw new Error(
          "Không tìm thấy word/document.xml. Vui lòng kiểm tra lại tệp DOCX.",
        );
      }

      const xml = await entry.async("string");
      const found = readSections(xml);

      setSections(found);

      setStatus(`Đã tải ${next.name} · Phát hiện ${found.length} phân đoạn.`);
    } catch (err) {
      setStatus(
        err instanceof Error ? err.message : "Không thể kiểm tra tài liệu.",
      );

      setSections([]);
    }
  }

  async function normalize() {
    if (!file) return;

    setBusy(true);
    setStatus("Đang xử lý tài liệu và chuẩn bị bản xem trước…");

    try {
      const result = await reconstructWithMargins(file, margins);

      // Mở lại gói ZIP vừa tạo để kiểm tra tính toàn vẹn cơ bản.
      const check = await JSZip.loadAsync(result.blob);

      if (
        !check.file("word/document.xml") ||
        !check.file("[Content_Types].xml")
      ) {
        throw new Error("Tệp đầu ra không vượt qua kiểm tra tính toàn vẹn.");
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
        "Đã tạo tài liệu sau chuẩn hóa. Vui lòng kiểm tra cả hai bản xem trước trước khi sử dụng. Phiên bản thử nghiệm chưa tự động kiểm tra các lỗi hiển thị.",
      );
    } catch (err) {
      setOutputBlob(null);

      setStatus(
        err instanceof Error ? err.message : "Không thể chuẩn hóa tài liệu.",
      );
    } finally {
      setBusy(false);
    }
  }

  function download() {
    if (!outputBlob || !file) return;

    const url = URL.createObjectURL(outputBlob);
    const a = document.createElement("a");

    a.href = url;
    a.download = `${file.name.replace(/\.docx$/i, "")}-da-chuan-hoa.docx`;
    a.click();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  const changed = outputBlob !== null;

  return (
    <main className="min-h-screen bg-[#f5f6f8] text-[#202735] [font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe_UI',sans-serif] [font-synthesis:none] [text-rendering:optimizeLegibility]">
      {/* Thanh tiêu đề */}
      <header className="flex h-[68px] items-center gap-3 border-b border-[#e7e9ee] bg-white px-[clamp(18px,4vw,60px)]">
        <div className="grid size-[34px] shrink-0 place-items-center rounded-[10px] bg-[#233fce] text-[19px] font-extrabold text-white shadow-[0_4px_10px_#233fce2b]">
          M
        </div>

        <div>
          <div className="text-[15px] font-extrabold tracking-[-0.3px]">
            Margin Lab
          </div>

          <div className="mt-[2px] text-[11px] text-[#8991a0]">
            Công cụ chuẩn hóa lề tài liệu DOCX
          </div>
        </div>

        <div className="ml-auto flex items-center gap-[7px] rounded-[30px] border border-[#dfeee6] bg-[#f4fbf6] px-[11px] py-[7px] text-[11px] text-[#34744c] max-[720px]:p-2 max-[720px]:text-[0px]">
          <span className="size-[6px] rounded-full bg-[#43a66b] max-[720px]:size-[7px]" />
          Xử lý trực tiếp trên thiết bị
        </div>
      </header>

      <div className="mx-auto max-w-[1500px] px-[clamp(18px,4vw,60px)] pb-6 pt-9 max-[720px]:pt-[25px]">
        {/* Giới thiệu */}
        <div className="mb-[27px] max-w-[680px]">
          <div className="mb-[10px] text-[10px] font-bold tracking-[1.7px] text-[#455fe1]">
            PHIÊN BẢN THỬ NGHIỆM · 0.1
          </div>

          <h1 className="m-0 text-[clamp(29px,3vw,42px)] leading-[1.13] font-extrabold tracking-[-1.8px] text-[#192235] max-[720px]:tracking-[-1px]">
            Chuẩn hóa lề văn bản.
            <br />
            <span className="text-[#5369dc]">
              Giữ nguyên nội dung tài liệu.
            </span>
          </h1>

          <p className="mt-[14px] mb-0 max-w-[640px] text-[13px] leading-[1.8] text-[#6f7786]">
            Điều chỉnh lề trang trong tài liệu DOCX mà vẫn giữ nguyên nội dung,
            hình ảnh, kiểu định dạng và các thành phần liên quan. So sánh tài
            liệu gốc với tài liệu sau khi chuẩn hóa.
          </p>
        </div>

        {/* Khu vực làm việc */}
        <div className="grid grid-cols-[310px_minmax(0,1fr)] items-start gap-5 max-[1050px]:grid-cols-[270px_minmax(0,1fr)] max-[1050px]:gap-[13px] max-[720px]:grid-cols-1">
          {/* Bảng điều khiển */}
          <aside className="rounded-[15px] border border-[#e5e8ee] bg-white p-[21px] shadow-[0_5px_20px_#18233b06] max-[1050px]:p-4 max-[720px]:order-1">
            {/* Tài liệu nguồn */}
            <div className="flex items-start gap-[11px]">
              <span className="grid size-[27px] shrink-0 place-items-center rounded-lg bg-[#eef1ff] text-[11px] font-bold text-[#4359d3]">
                01
              </span>

              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  Tài liệu nguồn
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  Chọn tài liệu Word cần xử lý.
                </p>
              </div>
            </div>

            {/* Tải tài liệu */}
            <label
              htmlFor="docx-file"
              className="mt-[18px] flex min-h-[142px] cursor-pointer flex-col items-center justify-center gap-[7px] rounded-[11px] border-[1.5px] border-dashed border-[#cdd3e3] bg-[#fafbff] p-[14px] text-center transition-colors duration-200 hover:border-[#6275e3] hover:bg-[#f6f7ff]"
            >
              <span className="grid size-[34px] place-items-center rounded-[10px] bg-[#e9edff] text-[23px] text-[#465bd2]">
                ↑
              </span>

              <strong className="max-w-full overflow-wrap-anywhere text-[12px]">
                {file ? file.name : "Kéo thả tệp DOCX vào đây"}
              </strong>

              <span className="text-[10px] text-[#8b93a1]">
                {file
                  ? `${(file.size / 1024).toFixed(0)} KB · Đã chọn`
                  : "Hoặc nhấn để chọn tệp · Chỉ hỗ trợ .docx"}
              </span>

              <input
                id="docx-file"
                type="file"
                accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                className="hidden"
                onChange={(e) => void loadFile(e.target.files?.[0])}
              />
            </label>

            {/* Thống kê phân đoạn */}
            {sections.length > 0 && (
              <div className="mt-3 flex items-center gap-[10px] rounded-[9px] bg-[#f6f8ff] p-[11px]">
                <div className="font-[Manrope] text-[20px] font-extrabold text-[#435bd3]">
                  {sections.length}
                </div>

                <div>
                  <strong className="block text-[11px]">
                    {sections.length === 1
                      ? "Phân đoạn được phát hiện"
                      : "Phân đoạn được phát hiện"}
                  </strong>

                  <span className="mt-[3px] block text-[10px] text-[#7e8796]">
                    Lề sẽ được áp dụng cho tất cả phân đoạn.
                  </span>
                </div>
              </div>
            )}

            <div className="my-[22px] h-px bg-[#edf0f4]" />

            {/* Thiết lập lề */}
            <div className="flex items-start gap-[11px]">
              <span className="grid size-[27px] shrink-0 place-items-center rounded-lg bg-[#eef1ff] text-[11px] font-bold text-[#4359d3]">
                02
              </span>

              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  Thiết lập lề trang
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  Đơn vị tính: milimét (mm).
                </p>
              </div>
            </div>

            <label
              htmlFor="profile"
              className="mt-[19px] mb-[7px] block text-[10px] font-bold text-[#737c8b]"
            >
              Mẫu thiết lập
            </label>

            <select
              id="profile"
              value={preset}
              onChange={(e) => choosePreset(e.target.value)}
              className="w-full rounded-lg border border-[#dfe3eb] bg-white p-[10px] text-[11px] text-[#333c4b] outline-none focus:border-[#687be3] focus:shadow-[0_0_0_3px_#5369dc16]"
            >
              <option value="vn-admin">
                Mẫu hành chính · 25 / 20 / 25 / 30
              </option>
              <option value="balanced">Cân đối · 25,4 mm mỗi cạnh</option>
              <option value="compact">Thu gọn · 15 mm mỗi cạnh</option>
              <option value="custom">Tùy chỉnh</option>
            </select>

            {/* Các thông số lề */}
            <div className="mt-[13px] grid grid-cols-2 gap-[10px]">
              {(["top", "right", "bottom", "left"] as const).map((key) => {
                const labels = {
                  top: "Lề trên",
                  right: "Lề phải",
                  bottom: "Lề dưới",
                  left: "Lề trái",
                };

                return (
                  <label key={key}>
                    <span className="mb-[6px] block text-[10px] text-[#737c8b]">
                      {labels[key]}
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
                );
              })}
            </div>

            {/* Chính sách bảo toàn dữ liệu */}
            <div className="mt-[19px] flex gap-[9px] rounded-[9px] border border-[#e1f1e7] bg-[#f3faf6] px-[10px] py-3">
              <span className="grid size-5 shrink-0 place-items-center rounded-full bg-[#d8f1e1] text-[12px] text-[#27814b]">
                ✓
              </span>

              <div>
                <strong className="text-[10px] text-[#2d6743]">
                  Nguyên tắc bảo toàn
                </strong>

                <p className="mt-1 mb-0 text-[10px] leading-[1.6] text-[#62806c]">
                  Giữ nguyên văn bản, bảng biểu, hình ảnh, kiểu định dạng và các
                  thành phần liên quan. Chỉ thay đổi thông số lề trang trong tài
                  liệu.
                </p>
              </div>
            </div>

            {/* Nút chuẩn hóa */}
            <button
              onClick={() => void normalize()}
              disabled={!file || busy || sections.length === 0}
              className="mt-[14px] flex w-full items-center justify-between rounded-[9px] border-0 bg-[#344fd0] px-[13px] py-3 text-[11px] font-semibold text-white shadow-[0_4px_10px_#344fd02b] transition-colors hover:bg-[#283fb9] disabled:cursor-not-allowed disabled:opacity-[0.45]"
            >
              {busy ? (
                <span className="flex items-center gap-2">
                  <span className="size-[13px] animate-spin rounded-full border-2 border-white/40 border-t-white" />
                  Đang xử lý…
                </span>
              ) : (
                <>
                  <span>Chuẩn hóa lề trang</span>
                  <span>→</span>
                </>
              )}
            </button>

            {/* Nút tải xuống */}
            {changed && (
              <button
                onClick={download}
                className="mt-[14px] flex w-full items-center justify-center gap-[9px] rounded-[9px] border border-[#d7deeb] bg-white px-[13px] py-3 text-[11px] font-semibold text-[#3449b5] hover:bg-[#f7f8ff]"
              >
                <span>↓</span>
                Tải tài liệu đã chuẩn hóa
              </button>
            )}

            {/* Trạng thái xử lý */}
            {status && (
              <div
                role="status"
                className="mt-[11px] overflow-wrap-anywhere rounded-[7px] bg-[#f7f8fa] px-[10px] py-[9px] text-[10px] leading-[1.6] text-[#747e8f]"
              >
                {status}
              </div>
            )}

            {/* Kết quả kiểm tra tính toàn vẹn */}
            {stats && (
              <div className="mt-[11px] rounded-[9px] border border-[#dcefe3] bg-[#f8fcf9] p-[11px]">
                <div className="flex items-center gap-[7px] text-[10px] text-[#2b7548]">
                  <span className="grid size-[17px] place-items-center rounded-full bg-[#d9f1e2]">
                    ✓
                  </span>

                  <strong>Kiểm tra tính toàn vẹn thành công</strong>
                </div>

                <div className="mt-3 grid grid-cols-2 gap-2">
                  <div className="rounded-[7px] border border-[#e6f1e9] bg-white p-2">
                    <strong className="block font-[Manrope] text-[18px] font-bold text-[#2d6844]">
                      {stats.sections}
                    </strong>

                    <span className="mt-[3px] block text-[9px] text-[#7e8d83]">
                      Phân đoạn đã cập nhật
                    </span>
                  </div>

                  <div className="rounded-[7px] border border-[#e6f1e9] bg-white p-2">
                    <strong className="block font-[Manrope] text-[18px] font-bold text-[#2d6844]">
                      {stats.preservedParts}
                    </strong>

                    <span className="mt-[3px] block text-[9px] text-[#7e8d83]">
                      Thành phần được giữ lại
                    </span>
                  </div>
                </div>

                <p className="mt-[9px] mb-0 text-[9px] leading-[1.6] text-[#73877a]">
                  Kết quả chỉ xác nhận tính toàn vẹn cơ bản của gói DOCX, không
                  đảm bảo tính chính xác về hiển thị. Vui lòng kiểm tra bản xem
                  trước để phát hiện các vấn đề liên quan đến bảng biểu hoặc
                  hình ảnh.
                </p>
              </div>
            )}

            {/* Thông tin bảo mật */}
            <div className="mt-[18px] flex items-center justify-center gap-[6px] text-[9px] text-[#959dab]">
              <span className="text-[16px]">⌁</span>
              Tài liệu được xử lý trực tiếp trên thiết bị, không tải lên máy
              chủ.
            </div>
          </aside>

          {/* Khu vực xem trước */}
          <section className="min-w-0 rounded-[15px] border border-[#e5e8ee] bg-white p-5 shadow-[0_5px_20px_#18233b06] max-[1050px]:p-[14px] max-[720px]:order-2">
            <div className="mb-[17px] flex items-center justify-between gap-[10px]">
              <div>
                <h2 className="mt-px mb-1 text-[14px] font-bold">
                  So sánh tài liệu
                </h2>

                <p className="m-0 text-[11px] text-[#8a92a0]">
                  Đối chiếu tài liệu gốc và tài liệu sau chuẩn hóa
                </p>
              </div>

              <div className="flex items-center gap-[6px] whitespace-nowrap text-[10px] text-[#737c8b]">
                <span className="ml-2 inline-block size-[7px] rounded-full bg-[#a6adb9]" />
                Bản gốc
                <span className="ml-[7px] inline-block size-[7px] rounded-full bg-[#4e65dc]" />
                Bản mới
              </div>
            </div>

            {/* Hai bản xem trước */}
            <div className="grid grid-cols-2 items-start gap-3 max-[1050px]:grid-cols-1 max-[720px]:grid-cols-1">
              <Preview
                blob={originalBlob}
                label="Tài liệu gốc"
                title={file?.name}
              />

              <Preview
                blob={outputBlob}
                label="Tài liệu sau chuẩn hóa"
                title={`${file?.name.replace(/\.docx$/i, "")}-da-chuan-hoa.docx`}
              />
            </div>

            {/* Lưu ý */}
            <div className="mt-[14px] flex gap-[9px] rounded-[9px] border border-[#f3e7bf] bg-[#fff9e9] p-3 text-[#92712b]">
              <span className="text-[15px]">ⓘ</span>

              <p className="m-0 text-[10px] leading-[1.7] text-[#8c7845]">
                <strong className="text-[#765d22]">
                  Lưu ý về phiên bản thử nghiệm:
                </strong>{" "}
                Bản xem trước chỉ phục vụ mục đích kiểm tra trực quan. Công cụ
                chưa tự động phát hiện nội dung bị tràn, chưa đối chiếu hình học
                bảng biểu và không đảm bảo Microsoft Word hay LibreOffice sẽ
                hiển thị tài liệu hoàn toàn giống nhau.
              </p>
            </div>
          </section>
        </div>

        {/* Chân trang */}
        <footer className="mt-5 flex justify-between gap-[10px] px-[3px] text-[9px] text-[#a0a6b1] max-[720px]:flex-col">
          <span>Margin Lab · Phiên bản thử nghiệm</span>
          <span>
            Chuẩn hóa cấu trúc DOCX · Chưa dành cho môi trường sản xuất
          </span>
        </footer>
      </div>
    </main>
  );
}
