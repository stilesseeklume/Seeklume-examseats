import React from "react";
import { FilePlus2, Info, Search, Trash2, Download, Sparkles } from "lucide-react";

export function HomeLanding({
  records,
  filteredRecords,
  recordFilter,
  setRecordFilter,
  clearLocalData,
  formatRecordTime,
  openRecord,
  exportRecord,
  removeRecord,
  onStart,
  onOpenAbout,
}) {
  return (
    <main className="app-shell home-shell">
      <header className="marketing-nav">
        <div>
          <strong>Seeklume ExamSeats</strong>
        </div>
        <button type="button" className="nav-link" onClick={onOpenAbout}><Info size={16} /> 隐私与开源</button>
      </header>
      <section className="home-hero">
        <div className="home-hero-copy fade-up">
          <span className="status-note"><Sparkles size={13} /> 不登录 · 不上传 · 不做云端同步</span>
          <h1>ExamSeats 排座工具</h1>
          <p className="home-sub">从成绩单到可打印考务材料，一次完成。导入物理类、历史类成绩单，自动排座、校验冲突、生成班主任表、考场信息表和门牌人数表。全程本机处理，学生数据只留在当前浏览器。</p>
          <div className="home-actions">
            <button type="button" className="primary" onClick={onStart}><FilePlus2 size={16} /> 开始排座</button>
          </div>
        </div>
      </section>
      {records.length > 0 && (
        <section className="local-records-panel" aria-label="本机历史记录">
          <div className="records-heading">
            <div>
              <h2>本机历史记录</h2>
              <p>只读取当前浏览器的本地记录。换设备、换浏览器或清理缓存后，这里可能看不到旧考试。</p>
            </div>
            <div className="record-toolbar">
              <label className="record-search">
                <Search size={15} />
                <input value={recordFilter} placeholder="搜索考试名称/日期/版本" onChange={(event) => setRecordFilter(event.target.value)} />
              </label>
              <button type="button" onClick={clearLocalData}><Trash2 size={16} /> 清空本机数据</button>
            </div>
          </div>
          {filteredRecords.length ? (
            <div className="record-list">
              {filteredRecords.slice(0, 8).map((record) => (
                <article className="record-card" key={record.id}>
                  <div>
                    <strong>{record.examName || "未命名考试"}</strong>
                    <span>{record.examDate || "未设置日期"} · 第 {record.version || 1} 版 · {formatRecordTime(record.createdAt)}</span>
                  </div>
                  <div className="record-actions">
                    <button type="button" onClick={() => openRecord(record)}>打开</button>
                    <button type="button" onClick={() => exportRecord(record)}><Download size={16} /> 导出</button>
                    <button type="button" onClick={() => removeRecord(record.id)}><Trash2 size={16} /> 删除</button>
                  </div>
                </article>
              ))}
            </div>
          ) : (
            <div className="empty local-empty">没有匹配的本机历史记录。</div>
          )}
        </section>
      )}
      <footer className="brand-footer">© Seeklume ExamSeats · 本机排座工具 · 数据不出浏览器</footer>
    </main>
  );
}
