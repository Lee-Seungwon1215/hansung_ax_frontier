import { useCallback, useMemo, useRef, useState } from "react";
import { googleLogout, useGoogleLogin } from "@react-oauth/google";
import "./App.css";

const API_BASE = "http://localhost:8001";
const MAX_EMAILS = 10;
const MAX_TEXT_LENGTH = 12000;
const MAX_FILE_BYTES = 8 * 1024 * 1024;

function decodeBase64Url(value = "") {
  try {
    const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
    const decoded = atob(normalized);
    return decodeURIComponent(
      decoded
        .split("")
        .map((char) => `%${`00${char.charCodeAt(0).toString(16)}`.slice(-2)}`)
        .join("")
    );
  } catch {
    try {
      return atob(value.replace(/-/g, "+").replace(/_/g, "/"));
    } catch {
      return "";
    }
  }
}

function stripHtml(html = "") {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return doc.body.textContent?.replace(/\s+/g, " ").trim() || "";
}

function findHeader(headers = [], name) {
  return headers.find((header) => header.name?.toLowerCase() === name.toLowerCase())?.value || "";
}

function collectMessageParts(part, list = []) {
  if (!part) return list;
  if (part.parts?.length) {
    part.parts.forEach((child) => collectMessageParts(child, list));
    return list;
  }
  list.push(part);
  return list;
}

function isTextLikeFile(name = "", mimeType = "") {
  return mimeType.startsWith("text/") || name.match(/\.(txt|csv|md|json)$/i);
}

function isDocumentFile(name = "", mimeType = "") {
  return (
    mimeType === "application/pdf" ||
    mimeType.includes("wordprocessingml") ||
    mimeType.includes("spreadsheetml") ||
    mimeType.includes("presentationml") ||
    name.match(/\.(pdf|docx|xlsx|pptx)$/i)
  );
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.split(",")[1] : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function App() {
  const fileInputRef = useRef(null);
  const [workInput, setWorkInput] = useState("");
  const [meetingFocus, setMeetingFocus] = useState("요약 보고서");
  const [draft, setDraft] = useState("");
  const [loading, setLoading] = useState(false);
  const [dataLoading, setDataLoading] = useState(false);
  const [dragActive, setDragActive] = useState(false);
  const [docCount, setDocCount] = useState(0);
  const [recentDocs, setRecentDocs] = useState([]);
  const [user, setUser] = useState(null);
  const [accessToken, setAccessToken] = useState(null);
  const [emails, setEmails] = useState([]);
  const [selectedEmailIds, setSelectedEmailIds] = useState([]);
  const [uploadedFiles, setUploadedFiles] = useState([]);

  const fetchAttachment = async (token, messageId, part) => {
    if (!part.body?.attachmentId) return null;

    const res = await fetch(
      `https://gmail.googleapis.com/gmail/v1/users/me/messages/${messageId}/attachments/${part.body.attachmentId}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await res.json();
    const mimeType = part.mimeType || "";
    const filename = part.filename || "첨부파일";
    const rawData = data.data || "";

    if (isTextLikeFile(filename, mimeType)) {
      return {
        name: filename,
        mimeType,
        text: decodeBase64Url(rawData).slice(0, MAX_TEXT_LENGTH),
      };
    }

    if (isDocumentFile(filename, mimeType)) {
      return {
        name: filename,
        mimeType,
        data: rawData,
      };
    }

    return {
      name: filename,
      mimeType,
      text: "(지원하지 않는 첨부 형식입니다. 파일명과 형식만 참고합니다.)",
    };
  };

  const fetchEmails = async (token) => {
    const since = new Date();
    since.setDate(since.getDate() - 14);
    const afterDate = `${since.getFullYear()}/${since.getMonth() + 1}/${since.getDate()}`;

    const listRes = await fetch(
      `https://gmail.googleapis.com/gmail/v1/users/me/messages?q=after:${afterDate}&maxResults=${MAX_EMAILS}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const listData = await listRes.json();
    if (!listData.messages) return [];

    return Promise.all(
      listData.messages.map(async (message) => {
        const msgRes = await fetch(
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${message.id}?format=full`,
          { headers: { Authorization: `Bearer ${token}` } }
        );
        const msg = await msgRes.json();
        const headers = msg.payload?.headers || [];
        const parts = collectMessageParts(msg.payload);
        const textParts = parts
          .filter((part) => part.body?.data && ["text/plain", "text/html"].includes(part.mimeType))
          .map((part) => {
            const raw = decodeBase64Url(part.body.data);
            return part.mimeType === "text/html" ? stripHtml(raw) : raw;
          })
          .filter(Boolean);

        const attachmentParts = parts.filter((part) => part.filename && part.body?.attachmentId);
        const attachments = await Promise.all(
          attachmentParts.slice(0, 5).map((part) => fetchAttachment(token, message.id, part))
        );

        return {
          id: msg.id,
          subject: findHeader(headers, "Subject") || "(제목 없음)",
          from: findHeader(headers, "From"),
          date: findHeader(headers, "Date"),
          body: (textParts.join("\n\n") || msg.snippet || "").slice(0, MAX_TEXT_LENGTH),
          attachments: attachments.filter(Boolean),
        };
      })
    );
  };

  const login = useGoogleLogin({
    scope: "openid profile email https://www.googleapis.com/auth/gmail.readonly",
    onSuccess: async (tokenResponse) => {
      setDataLoading(true);
      const token = tokenResponse.access_token;
      setAccessToken(token);

      try {
        const profileRes = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
          headers: { Authorization: `Bearer ${token}` },
        });
        setUser(await profileRes.json());
        const loadedEmails = await fetchEmails(token);
        setEmails(loadedEmails);
        setSelectedEmailIds(loadedEmails.map((email) => email.id));
      } catch {
        alert("Gmail 데이터를 불러오지 못했습니다.");
      } finally {
        setDataLoading(false);
      }
    },
    onError: () => alert("Google 로그인에 실패했습니다."),
  });

  const handleLogout = () => {
    googleLogout();
    setUser(null);
    setAccessToken(null);
    setEmails([]);
    setSelectedEmailIds([]);
  };

  const toggleEmail = (emailId) => {
    setSelectedEmailIds((prev) =>
      prev.includes(emailId) ? prev.filter((id) => id !== emailId) : [...prev, emailId]
    );
  };

  const selectAllEmails = () => setSelectedEmailIds(emails.map((email) => email.id));
  const clearEmailSelection = () => setSelectedEmailIds([]);

  const readFiles = useCallback(async (files) => {
    const nextFiles = await Promise.all(
      Array.from(files).map(async (file) => {
        if (file.size > MAX_FILE_BYTES) {
          return {
            id: `${file.name}-${file.lastModified}`,
            name: file.name,
            type: file.type || "unknown",
            size: file.size,
            text: "(8MB를 초과해 본문 추출 대상에서 제외했습니다.)",
          };
        }

        if (isTextLikeFile(file.name, file.type)) {
          return {
            id: `${file.name}-${file.lastModified}`,
            name: file.name,
            type: file.type || "text/plain",
            size: file.size,
            text: (await file.text()).slice(0, MAX_TEXT_LENGTH),
          };
        }

        if (isDocumentFile(file.name, file.type)) {
          return {
            id: `${file.name}-${file.lastModified}`,
            name: file.name,
            type: file.type || "unknown",
            size: file.size,
            data: await fileToBase64(file),
          };
        }

        return {
          id: `${file.name}-${file.lastModified}`,
          name: file.name,
          type: file.type || "unknown",
          size: file.size,
          text: "(지원하지 않는 첨부 형식입니다. 파일명과 형식만 참고합니다.)",
        };
      })
    );
    setUploadedFiles((prev) => [...nextFiles, ...prev].slice(0, 12));
  }, []);

  const selectedEmails = useMemo(
    () => emails.filter((email) => selectedEmailIds.includes(email.id)),
    [emails, selectedEmailIds]
  );

  const emailSources = useMemo(() => {
    return selectedEmails.map((email) => ({
      subject: email.subject,
      from: email.from,
      date: email.date,
      body: email.body,
      attachments: email.attachments,
    }));
  }, [selectedEmails]);

  const generateReport = async () => {
    if (!workInput.trim() && !selectedEmails.length && !uploadedFiles.length) {
      alert("보고서에 반영할 요청, 선택한 Gmail, 첨부파일 중 하나 이상이 필요합니다.");
      return;
    }

    setLoading(true);
    setDraft("");
    try {
      const res = await fetch(`${API_BASE}/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          work_content: workInput,
          target_committee: "자료 기반 보고서",
          meeting_focus: meetingFocus,
          email_sources: emailSources,
          uploaded_files: uploadedFiles.map(({ name, type, text, data }) => ({ name, type, text, data })),
        }),
      });
      const data = await res.json();
      setDraft(data.draft || "보고서 초안을 생성하지 못했습니다.");
      setDocCount((prev) => prev + 1);
      const today = new Date().toLocaleDateString("ko-KR");
      setRecentDocs((prev) =>
        [{ title: `AI 보고서 초안 (${today})`, status: "초안 생성" }, ...prev].slice(0, 5)
      );
    } catch {
      setDraft("서버 연결에 실패했습니다. FastAPI 백엔드가 실행 중인지 확인해 주세요.");
    } finally {
      setLoading(false);
    }
  };

  const downloadFile = async (endpoint, filename) => {
    const res = await fetch(`${API_BASE}/${endpoint}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ content: draft, filename }),
    });
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  return (
    <main className="app-shell">
      <section className="topbar">
        <div>
          <p className="eyebrow">Hansung AI Report</p>
          <h1>자료 기반 AI 보고서 변환</h1>
        </div>
        {user ? (
          <div className="user-box">
            {user.picture && <img src={user.picture} alt="프로필" />}
            <span>{user.name}</span>
            <button type="button" onClick={handleLogout}>로그아웃</button>
          </div>
        ) : (
          <button type="button" className="primary-light" onClick={() => login()}>
            Gmail 연결
          </button>
        )}
      </section>

      <section className="summary-grid">
        <div>
          <span>생성 문서</span>
          <strong>{docCount}건</strong>
        </div>
        <div>
          <span>선택한 메일</span>
          <strong>{selectedEmails.length}/{emails.length}건</strong>
        </div>
        <div>
          <span>첨부 자료</span>
          <strong>{uploadedFiles.length}개</strong>
        </div>
      </section>

      <section className="workspace">
        <div className="panel editor-panel">
          <div className="section-title">
            <div>
              <p className="eyebrow">Report Builder</p>
              <h2>보고서 생성 정보</h2>
            </div>
            <select value={meetingFocus} onChange={(event) => setMeetingFocus(event.target.value)}>
              <option>요약 보고서</option>
              <option>업무 정리 보고서</option>
              <option>회의 보고서</option>
              <option>PPT 발표용 보고서</option>
              <option>검토 의견 보고서</option>
            </select>
          </div>

          <label className="field-label" htmlFor="work-input">요청 내용</label>
          <textarea
            id="work-input"
            value={workInput}
            onChange={(event) => setWorkInput(event.target.value)}
            placeholder={"예시:\n제안서 정리\n회의록 정리\n심의의결서 정리\n결재서 정리\n보고서 정리\n선택한 메일과 첨부파일을 참고해서 보고서 형식으로 요약해줘."}
          />

          <div
            className={`dropzone ${dragActive ? "is-active" : ""}`}
            onDragOver={(event) => {
              event.preventDefault();
              setDragActive(true);
            }}
            onDragLeave={() => setDragActive(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDragActive(false);
              readFiles(event.dataTransfer.files);
            }}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              multiple
              onChange={(event) => readFiles(event.target.files)}
            />
            <strong>파일을 드래그하거나 클릭해서 첨부</strong>
            <span>PDF, Word, Excel, PPT, txt, csv, md, json 내용을 보고서 근거로 활용합니다.</span>
          </div>

          {uploadedFiles.length > 0 && (
            <div className="file-list">
              {uploadedFiles.map((file) => (
                <div key={file.id || file.name}>
                  <span>{file.name}</span>
                  <button
                    type="button"
                    onClick={() => setUploadedFiles((prev) => prev.filter((item) => item !== file))}
                  >
                    제거
                  </button>
                </div>
              ))}
            </div>
          )}

          <button type="button" className="generate-button" onClick={generateReport} disabled={loading}>
            {loading ? "AI가 보고서 초안을 작성 중입니다..." : "보고서 초안 생성"}
          </button>
        </div>

        <aside className="panel source-panel">
          <div className="section-title compact">
            <div>
              <p className="eyebrow">Sources</p>
              <h2>Gmail 참고자료</h2>
            </div>
            {dataLoading && <span className="loading-pill">불러오는 중</span>}
          </div>

          {!accessToken && <p className="empty-text">Gmail을 연결하면 최근 14일 메일을 불러오고, 보고서에 반영할 메일을 직접 선택할 수 있습니다.</p>}
          {accessToken && emails.length === 0 && !dataLoading && <p className="empty-text">최근 메일이 없습니다.</p>}

          {emails.length > 0 && (
            <div className="mail-toolbar">
              <span>{selectedEmails.length}개 선택됨</span>
              <div>
                <button type="button" onClick={selectAllEmails}>전체 선택</button>
                <button type="button" onClick={clearEmailSelection}>선택 해제</button>
              </div>
            </div>
          )}

          {emails.map((email) => {
            const checked = selectedEmailIds.includes(email.id);
            return (
              <article className={`mail-item ${checked ? "is-selected" : ""}`} key={email.id}>
                <label className="mail-check">
                  <input type="checkbox" checked={checked} onChange={() => toggleEmail(email.id)} />
                  <span>
                    <strong>{email.subject}</strong>
                    <small>{email.from}</small>
                  </span>
                </label>
                <p>{email.body || "본문 없음"}</p>
                {email.attachments.length > 0 && (
                  <small>첨부 {email.attachments.map((file) => file.name).join(", ")}</small>
                )}
              </article>
            );
          })}

          <div className="recent-box">
            <h3>최근 생성 문서</h3>
            {recentDocs.length === 0 ? (
              <p className="empty-text">아직 생성된 문서가 없습니다.</p>
            ) : (
              recentDocs.map((doc) => (
                <div key={doc.title}>
                  <span>{doc.title}</span>
                  <b>{doc.status}</b>
                </div>
              ))
            )}
          </div>
        </aside>
      </section>

      {draft && (
        <section className="panel draft-panel">
          <div className="section-title">
            <div>
              <p className="eyebrow">Draft</p>
              <h2>AI 생성 보고서 초안</h2>
            </div>
            <div className="download-actions">
              <button type="button" onClick={() => downloadFile("save-word", "AI_보고서_초안.docx")}>
                Word 보고서 저장
              </button>
              <button type="button" onClick={() => downloadFile("save-ppt", "AI_보고서_초안.pptx")}>
                PPT 보고서 저장
              </button>
            </div>
          </div>
          <textarea value={draft} onChange={(event) => setDraft(event.target.value)} />
        </section>
      )}
    </main>
  );
}

export default App;
