import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { googleLogout, useGoogleLogin } from "@react-oauth/google";
import "./App.css";

const API_BASE = process.env.REACT_APP_API_BASE || "http://127.0.0.1:8001";
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

function StatusBadge({ status }) {
  return <span className={`status-badge status-${status}`}>{status}</span>;
}

function App() {
  const fileInputRef = useRef(null);
  const approvalFileInputRef = useRef(null);
  const [activeView, setActiveView] = useState("intranet");
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
  const [selectedApprovalIds, setSelectedApprovalIds] = useState([]);
  const [showAllApprovals, setShowAllApprovals] = useState(false);
  const [showAllNotices, setShowAllNotices] = useState(false);
  const [showAllDocuments, setShowAllDocuments] = useState(false);
  const [showAllContacts, setShowAllContacts] = useState(false);
  const [intranet, setIntranet] = useState(null);
  const [savedReports, setSavedReports] = useState([]);
  const [intranetLoading, setIntranetLoading] = useState(false);
  const [approvalForm, setApprovalForm] = useState({
    title: "",
    category: "심의/의결서",
    requester: "SW중심대학사업단 운영위원회",
    department: "SW중심대학사업단",
    amount: "",
    content: "",
  });
  const [approvalFiles, setApprovalFiles] = useState([]);

  const loadIntranet = useCallback(async () => {
    setIntranetLoading(true);
    try {
      let lastError = null;
      for (let attempt = 0; attempt < 45; attempt += 1) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 2500);
        try {
          const res = await fetch(`${API_BASE}/intranet/dashboard`, {
            signal: controller.signal,
          });
          if (!res.ok) {
            throw new Error(`Dashboard request failed: ${res.status}`);
          }
          setIntranet(await res.json());
          return;
        } catch (error) {
          lastError = error;
        } finally {
          clearTimeout(timeoutId);
        }
        if (attempt < 44) {
          await new Promise((resolve) => setTimeout(resolve, 1000));
        }
      }
      throw lastError;
    } catch {
      setIntranet((prev) => prev);
    } finally {
      setIntranetLoading(false);
    }
  }, []);

  const loadSavedReports = useCallback(async () => {
    try {
      const res = await fetch(`${API_BASE}/weekly-reports`);
      const data = await res.json();
      setSavedReports(data.reports || []);
    } catch {
      setSavedReports([]);
    }
  }, []);

  useEffect(() => {
    loadIntranet();
    loadSavedReports();
  }, [loadIntranet, loadSavedReports]);

  useEffect(() => {
    if (intranet?.approvals?.length) {
      setSelectedApprovalIds(intranet.approvals.map((approval) => approval.id));
    }
  }, [intranet]);

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
      return { name: filename, mimeType, text: decodeBase64Url(rawData).slice(0, MAX_TEXT_LENGTH) };
    }
    if (isDocumentFile(filename, mimeType)) {
      return { name: filename, mimeType, data: rawData };
    }
    return { name: filename, mimeType, text: "지원하지 않는 첨부 형식입니다." };
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

  const selectedEmails = useMemo(
    () => emails.filter((email) => selectedEmailIds.includes(email.id)),
    [emails, selectedEmailIds]
  );

  const emailSources = useMemo(
    () =>
      selectedEmails.map((email) => ({
        subject: email.subject,
        from: email.from,
        date: email.date,
        body: email.body,
        attachments: email.attachments,
      })),
    [selectedEmails]
  );

  const toggleEmail = (emailId) => {
    setSelectedEmailIds((prev) =>
      prev.includes(emailId) ? prev.filter((id) => id !== emailId) : [...prev, emailId]
    );
  };

  const readFiles = useCallback(async (files) => {
    const nextFiles = await Promise.all(
      Array.from(files).map(async (file) => {
        if (file.size > MAX_FILE_BYTES) {
          return {
            id: `${file.name}-${file.lastModified}`,
            name: file.name,
            type: file.type || "unknown",
            size: file.size,
            text: "8MB를 초과해 본문 추출 대상에서 제외했습니다.",
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
          text: "지원하지 않는 첨부 형식입니다.",
        };
      })
    );
    setUploadedFiles((prev) => [...nextFiles, ...prev].slice(0, 12));
  }, []);

  const readApprovalFiles = useCallback(async (files) => {
    const nextFiles = await Promise.all(
      Array.from(files).map(async (file) => {
        const base = {
          id: `${file.name}-${file.lastModified}`,
          name: file.name,
          type: file.type || "unknown",
          size: file.size,
        };
        if (file.size > MAX_FILE_BYTES) {
          return { ...base, text: "8MB를 초과해 본문 추출 대상에서 제외했습니다." };
        }
        if (isTextLikeFile(file.name, file.type)) {
          return { ...base, text: (await file.text()).slice(0, MAX_TEXT_LENGTH) };
        }
        if (isDocumentFile(file.name, file.type)) {
          return { ...base, data: await fileToBase64(file) };
        }
        return { ...base, text: "지원하지 않는 첨부 형식입니다." };
      })
    );
    setApprovalFiles((prev) => [...nextFiles, ...prev].slice(0, 8));
  }, []);

  const generateReport = async () => {
    if (!workInput.trim() && !selectedEmails.length && !uploadedFiles.length) {
      alert("보고서에 반영할 요청, Gmail, 첨부파일 중 하나 이상이 필요합니다.");
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
      setRecentDocs((prev) => [{ title: `AI 보고서 초안 (${today})`, status: "초안 생성" }, ...prev].slice(0, 5));
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

  const downloadTemplate = async (doc) => {
    const res = await fetch(`${API_BASE}/documents/${doc.id}/download`);
    const contentType = res.headers.get("content-type") || "";
    if (!res.ok || !contentType.includes("wordprocessingml")) {
      alert("양식 파일을 다운로드하지 못했습니다. 백엔드 서버를 재시작한 뒤 다시 시도해 주세요.");
      return;
    }
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${doc.title}.docx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const submitApproval = async (event) => {
    event.preventDefault();
    if (!approvalForm.title.trim() || !approvalForm.content.trim()) {
      alert("기안 제목과 내용을 입력해 주세요.");
      return;
    }
    const res = await fetch(`${API_BASE}/intranet/approvals`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ...approvalForm,
        attachments: approvalFiles.map(({ name, type, text, data }) => ({ name, type, text, data })),
      }),
    });
    const newApproval = await res.json();
    setIntranet((prev) => ({
      ...prev,
      approvals: [newApproval, ...(prev?.approvals || [])],
      summary: {
        ...(prev?.summary || {}),
        waitingApprovals: (prev?.summary?.waitingApprovals || 0) + 1,
      },
    }));
    setApprovalForm((prev) => ({ ...prev, title: "", amount: "", content: "" }));
    setApprovalFiles([]);
  };

  const saveWeeklyReport = async (status) => {
    if (!draft.trim()) {
      alert("저장할 보고서 초안이 없습니다.");
      return;
    }
    const title = draft.match(/^#\s+(.+)$/m)?.[1] || "SW중심대학사업단 운영위원회 주간업무보고";
    const res = await fetch(`${API_BASE}/weekly-reports`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        title,
        period: "이번 주",
        source_approval_ids: selectedApprovalIds,
        draft_content: draft,
        final_content: draft,
        status,
        created_by: intranet?.currentUser?.name || "운영위원회 직원",
      }),
    });
    const report = await res.json();
    setSavedReports((prev) => [report, ...prev]);
    alert(`${status}으로 저장했습니다.`);
  };

  const openSavedReport = (report) => {
    setDraft(report.final_content || report.draft_content || "");
    setActiveView("report");
  };

  const approvals = intranet?.approvals || [];
  const visibleApprovals = showAllApprovals ? approvals : approvals.slice(0, 3);
  const selectedApprovals = approvals.filter((approval) => selectedApprovalIds.includes(approval.id));
  const notices = intranet?.notices || [];
  const documents = intranet?.documents || [];
  const contacts = intranet?.contacts || [];
  const visibleNotices = showAllNotices ? notices : notices.slice(0, 2);
  const visibleDocuments = showAllDocuments ? documents : documents.slice(0, 2);
  const visibleContacts = showAllContacts ? contacts : contacts.slice(0, 2);
  const summary = intranet?.summary || {
    waitingApprovals: 0,
    activeApprovals: 0,
    completedApprovals: 0,
    notices: 0,
  };

  const toggleApproval = (approvalId) => {
    setSelectedApprovalIds((prev) =>
      prev.includes(approvalId) ? prev.filter((id) => id !== approvalId) : [...prev, approvalId]
    );
  };

  const prepareWeeklyReport = async () => {
    if (!selectedApprovals.length) {
      alert("주간보고서에 반영할 결재건을 선택해 주세요.");
      return;
    }

    setMeetingFocus("업무 정리 보고서");
    try {
      const res = await fetch(`${API_BASE}/weekly-reports/compose-request`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ approval_ids: selectedApprovalIds }),
      });
      const data = await res.json();
      setWorkInput(data.work_content || "");
    } catch {
      alert("주간보고서 요청내용을 생성하지 못했습니다.");
      return;
    }
    setActiveView("report");
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-block">
          <img className="hansung-logo" src="/hansung-logo.svg" alt="한성대학교" />
          <div>
            <p className="brand-wordmark">HANSUNG UNIVERSITY</p>
            <h1>SW중심대학 사업단 운영위원회</h1>
          </div>
        </div>
        <nav className="view-tabs" aria-label="주요 화면">
          <button className={activeView === "intranet" ? "is-active" : ""} onClick={() => setActiveView("intranet")}>
            인트라넷
          </button>
          <button className={activeView === "report" ? "is-active" : ""} onClick={() => setActiveView("report")}>
            AI 보고서
          </button>
        </nav>
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
      </header>

      {activeView === "intranet" ? (
        <section className="intranet-view">
          <section className="summary-grid">
            <div>
              <span>결재 대기</span>
              <strong>{summary.waitingApprovals}건</strong>
            </div>
            <div>
              <span>진행 중 결재</span>
              <strong>{summary.activeApprovals}건</strong>
            </div>
            <div>
              <span>완료 결재</span>
              <strong>{summary.completedApprovals}건</strong>
            </div>
            <div>
              <span>공지사항</span>
              <strong>{summary.notices}건</strong>
            </div>
          </section>

          {intranetLoading && <p className="empty-text">인트라넷 더미 데이터를 불러오는 중입니다.</p>}

          <section className="intranet-grid">
            <div className="panel approval-panel">
              <div className="section-title">
                <div>
                  <p className="eyebrow">Electronic Approval</p>
                  <h2>운영위원회 전자결재함</h2>
                </div>
                <button type="button" className="secondary-button" onClick={prepareWeeklyReport}>
                  선택 건으로 주간보고서 작성
                </button>
              </div>
              <div className="mail-toolbar approval-toolbar">
                <span>이번 주 처리 업무 {selectedApprovals.length}/{approvals.length}건 선택</span>
                <div>
                  <button type="button" onClick={() => setSelectedApprovalIds(approvals.map((approval) => approval.id))}>전체 선택</button>
                  <button type="button" onClick={() => setSelectedApprovalIds([])}>선택 해제</button>
                </div>
              </div>
              <div className="approval-list">
                {visibleApprovals.map((approval) => (
                  <article key={approval.id} className="approval-item">
                    <label className="approval-check">
                      <input
                        type="checkbox"
                        checked={selectedApprovalIds.includes(approval.id)}
                        onChange={() => toggleApproval(approval.id)}
                      />
                      <span>
                        <strong>{approval.title}</strong>
                        <small>
                          {approval.id} · {approval.department} · {approval.requester} · {approval.submittedAt}
                        </small>
                      </span>
                    </label>
                    <div className="approval-body">
                      <p>{approval.content}</p>
                      <small>
                        {approval.workType || "업무분류 확인 필요"} · {approval.outputType || approval.category}
                      </small>
                      <small>
                        첨부 {approval.attachments?.length || 0}개
                        {approval.attachments?.length ? ` · ${approval.attachments.map((file) => file.name).join(", ")}` : ""}
                      </small>
                      <small>후속조치: {approval.followUp || "확인 필요"}</small>
                    </div>
                    <div className="approval-meta">
                      <StatusBadge status={approval.status} />
                      <span>{approval.category}</span>
                      <b>{approval.currentApprover}</b>
                    </div>
                  </article>
                ))}
              </div>
              {approvals.length > 3 && (
                <button
                  type="button"
                  className="more-button"
                  onClick={() => setShowAllApprovals((prev) => !prev)}
                >
                  {showAllApprovals ? "접기" : `더보기 ${approvals.length - 3}건`}
                </button>
              )}
            </div>

            <aside className="panel draft-approval">
              <div className="section-title compact">
                <div>
                  <p className="eyebrow">Draft</p>
                  <h2>결재 기안 작성</h2>
                </div>
              </div>
              <form onSubmit={submitApproval}>
                <label className="field-label">기안 제목</label>
                <input
                  value={approvalForm.title}
                  onChange={(event) => setApprovalForm((prev) => ({ ...prev, title: event.target.value }))}
                  placeholder="예: 2026년 SW중심대학사업 예산 변경 심의"
                />
                <label className="field-label">문서 구분</label>
                <select
                  value={approvalForm.category}
                  onChange={(event) => setApprovalForm((prev) => ({ ...prev, category: event.target.value }))}
                >
                  <option>심의/의결서</option>
                  <option>제안서</option>
                  <option>회의록</option>
                  <option>보고서</option>
                  <option>결재서</option>
                </select>
                <label className="field-label">기안 내용</label>
                <textarea
                  value={approvalForm.content}
                  onChange={(event) => setApprovalForm((prev) => ({ ...prev, content: event.target.value }))}
                  placeholder="사업계획, 예산 조정, 운영 규정, 교육과정 개편, 산학협력, KPI 점검 내용을 입력하세요."
                />
                <label className="field-label">첨부 서류</label>
                <div className="mini-dropzone" onClick={() => approvalFileInputRef.current?.click()}>
                  <input
                    ref={approvalFileInputRef}
                    type="file"
                    multiple
                    onChange={(event) => readApprovalFiles(event.target.files)}
                  />
                  <strong>제안서, 회의록, 심의/의결서, 보고서, 결재서 첨부</strong>
                  <span>시연용 결재에는 파일명과 형식이 저장됩니다.</span>
                </div>
                {approvalFiles.length > 0 && (
                  <div className="file-list compact">
                    {approvalFiles.map((file) => (
                      <div key={file.id}>
                        <span>{file.name}</span>
                        <button
                          type="button"
                          onClick={() => setApprovalFiles((prev) => prev.filter((item) => item !== file))}
                        >
                          제거
                        </button>
                      </div>
                    ))}
                  </div>
                )}
                <button type="submit" className="generate-button">결재 상신</button>
              </form>
            </aside>
          </section>

          <section className="intranet-grid lower-grid">
            <div className="panel">
              <div className="section-title compact">
                <div>
                  <p className="eyebrow">Notice</p>
                  <h2>공지사항</h2>
                </div>
              </div>
              <div className="notice-list">
                {visibleNotices.map((notice) => (
                  <article key={notice.id}>
                    <span>{notice.priority}</span>
                    <strong>{notice.title}</strong>
                    <p>{notice.summary}</p>
                    <small>{notice.department} · {notice.date}</small>
                  </article>
                ))}
              </div>
              {notices.length > 2 && (
                <button type="button" className="more-button compact-more" onClick={() => setShowAllNotices((prev) => !prev)}>
                  {showAllNotices ? "접기" : `더보기 ${notices.length - 2}건`}
                </button>
              )}
            </div>

            <div className="panel">
              <div className="section-title compact">
                <div>
                  <p className="eyebrow">Weekly Report</p>
                  <h2>주간보고서 양식</h2>
                </div>
              </div>
              <div className="resource-list">
                {visibleDocuments.map((doc) => (
                  <div key={doc.id}>
                    <span>{doc.type.toUpperCase()}</span>
                    <strong>{doc.title}</strong>
                    <small>{doc.owner} · {doc.updatedAt}</small>
                    <button type="button" className="template-download-button" onClick={() => downloadTemplate(doc)}>
                      양식 다운로드
                    </button>
                  </div>
                ))}
              </div>
              {documents.length > 2 && (
                <button type="button" className="more-button compact-more" onClick={() => setShowAllDocuments((prev) => !prev)}>
                  {showAllDocuments ? "접기" : `더보기 ${documents.length - 2}건`}
                </button>
              )}
            </div>

            <div className="panel">
              <div className="section-title compact">
                <div>
                  <p className="eyebrow">Directory</p>
                  <h2>사업단 담당 연락처</h2>
                </div>
              </div>
              <div className="contact-list">
                {visibleContacts.map((contact) => (
                  <div key={contact.email}>
                    <strong>{contact.name}</strong>
                    <span>{contact.department} · {contact.role}</span>
                    <small>{contact.email} · {contact.phone}</small>
                  </div>
                ))}
              </div>
              {contacts.length > 2 && (
                <button type="button" className="more-button compact-more" onClick={() => setShowAllContacts((prev) => !prev)}>
                  {showAllContacts ? "접기" : `더보기 ${contacts.length - 2}건`}
                </button>
              )}
            </div>
          </section>
        </section>
      ) : (
        <>
          <section className="report-status-strip">
            <div className="report-status-item">
              <span>생성 문서</span>
              <strong>{docCount}건</strong>
            </div>
            <div className="report-status-item">
              <span>선택 메일</span>
              <strong>{selectedEmails.length}/{emails.length}건</strong>
            </div>
            <div className="report-status-item">
              <span>첨부 자료</span>
              <strong>{uploadedFiles.length}개</strong>
            </div>
          </section>

          <section className="workspace report-three-column">
            <div className="panel editor-panel report-tool-card">
              <div className="section-title">
                <div>
                  <p className="eyebrow">Step 1</p>
                  <h2>자료와 요청</h2>
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
                placeholder={"예시:\n선택한 메일과 첨부파일을 참고해서 운영위원회 심의/의결서 초안을 작성해줘.\nKPI 달성 현황을 보고서 형식으로 정리해줘.\n예산 변경 사유를 결재서 형식으로 정리해줘."}
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
                <input ref={fileInputRef} type="file" multiple onChange={(event) => readFiles(event.target.files)} />
                <strong>파일을 드래그하거나 클릭해서 첨부</strong>
                <span>PDF, Word, Excel, PPT, txt, csv, md, json을 보고서 근거로 사용합니다.</span>
              </div>

              {uploadedFiles.length > 0 && (
                <div className="file-list">
                  {uploadedFiles.map((file) => (
                    <div key={file.id || file.name}>
                      <span>{file.name}</span>
                      <button type="button" onClick={() => setUploadedFiles((prev) => prev.filter((item) => item !== file))}>
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

            {draft ? (
              <section className="panel draft-panel inline-draft-panel report-tool-card draft-card">
                <div className="section-title">
                  <div>
                    <p className="eyebrow">Step 2</p>
                    <h2>초안 수정</h2>
                  </div>
                  <div className="download-actions">
                    <button type="button" onClick={() => saveWeeklyReport("초안")}>초안 저장</button>
                    <button type="button" onClick={() => saveWeeklyReport("최종")}>최종본 저장</button>
                    <button type="button" onClick={() => downloadFile("save-word", "AI_보고서_초안.docx")}>Word 저장</button>
                    <button type="button" onClick={() => downloadFile("save-ppt", "AI_보고서_초안.pptx")}>PPT 저장</button>
                  </div>
                </div>
                <textarea
                  value={draft}
                  readOnly={false}
                  spellCheck={false}
                  onChange={(event) => setDraft(event.currentTarget.value)}
                  onInput={(event) => setDraft(event.currentTarget.value)}
                  onCompositionEnd={(event) => setDraft(event.currentTarget.value)}
                />
              </section>
            ) : (
              <section className="panel draft-panel inline-draft-panel empty-draft-panel report-tool-card draft-card">
                <div className="section-title">
                  <div>
                    <p className="eyebrow">Step 2</p>
                    <h2>초안 수정</h2>
                  </div>
                </div>
                <p className="empty-text">보고서 초안을 생성하면 이 영역에서 바로 확인하고 수정할 수 있습니다.</p>
              </section>
            )}

            <div className="report-side-stack">
              <aside className="panel source-panel report-tool-card source-card">
                <div className="section-title compact">
                  <div>
                    <p className="eyebrow">Step 3</p>
                    <h2>참고자료</h2>
                  </div>
                  {dataLoading && <span className="loading-pill">불러오는 중</span>}
                </div>

                {!accessToken && <p className="empty-text">Gmail을 연결하면 최근 14일 메일을 불러오고 보고서에 반영할 메일을 선택할 수 있습니다.</p>}
                {accessToken && emails.length === 0 && !dataLoading && <p className="empty-text">최근 메일이 없습니다.</p>}

                {emails.length > 0 && (
                  <div className="mail-toolbar">
                    <span>{selectedEmails.length}개 선택됨</span>
                    <div>
                      <button type="button" onClick={() => setSelectedEmailIds(emails.map((email) => email.id))}>전체 선택</button>
                      <button type="button" onClick={() => setSelectedEmailIds([])}>선택 해제</button>
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
                      {email.attachments.length > 0 && <small>첨부 {email.attachments.map((file) => file.name).join(", ")}</small>}
                    </article>
                  );
                })}

                <div className="recent-box">
                  <h3>저장된 주간보고서</h3>
                  {savedReports.length === 0 && recentDocs.length === 0 ? (
                    <p className="empty-text">아직 저장한 보고서가 없습니다.</p>
                  ) : (
                    [...savedReports, ...recentDocs].slice(0, 5).map((doc) => (
                      <div key={doc.id || doc.title}>
                        <button
                          type="button"
                          className="saved-report-button"
                          onClick={() => doc.id && openSavedReport(doc)}
                          disabled={!doc.id}
                        >
                          {doc.title}
                        </button>
                        <b>{doc.status}</b>
                      </div>
                    ))
                  )}
                </div>
              </aside>
            </div>
          </section>
        </>
      )}
    </main>
  );
}

export default App;
