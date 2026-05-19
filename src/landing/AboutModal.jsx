import React from "react";

export function AboutModal({ onClose }) {
  return (
    <div className="fullscreen-overlay about-overlay" role="dialog" aria-modal="true" aria-label="隐私与开源说明">
      <section className="about-card">
        <div className="about-head">
          <div>
            <span>Seeklume ExamSeats</span>
            <h2>隐私与开源</h2>
          </div>
          <button type="button" onClick={onClose}>关闭</button>
        </div>
        <div className="about-grid">
          <article>
            <strong>数据留在本机</strong>
            <p>本工具不登录、不上传、不做云端同步。学生名单、考场配置、考试时间和本机历史只保存在当前浏览器。</p>
          </article>
          <article>
            <strong>历史记录边界</strong>
            <p>更换电脑、清理浏览器缓存、使用无痕模式或换域名访问，都可能导致历史不可见。长期备份请导出 Excel 自行保存。</p>
          </article>
          <article>
            <strong>开源许可</strong>
            <p>项目代码按 MIT License 开源发布，可自由使用、复制、修改和分发；请保留许可证与必要署名。</p>
          </article>
        </div>
        <p className="about-note">数据只在本机，导出后由你自己掌握。</p>
      </section>
    </div>
  );
}
