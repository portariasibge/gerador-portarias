document.addEventListener("DOMContentLoaded", () => {
  // ====== ABAS (bate com o seu index.html) ======
  const tabButtons = document.querySelectorAll(".tab-button");
  const tabContents = document.querySelectorAll(".tab-content");

  function showTab(tabId) {
    tabButtons.forEach(btn => btn.classList.remove("active"));
    tabContents.forEach(div => (div.style.display = "none"));

    const activeBtn = document.querySelector(`.tab-button[data-tab="${tabId}"]`);
    const activeContent = document.getElementById(tabId);

    if (activeBtn) activeBtn.classList.add("active");
    if (activeContent) activeContent.style.display = "block";
  }

  tabButtons.forEach(btn => {
    btn.addEventListener("click", () => showTab(btn.dataset.tab));
  });

  // abre a primeira aba por padrão
  if (tabButtons.length) showTab(tabButtons[0].dataset.tab);

  // ====== SUBMISSÃO (mapeando forms -> rotas Flask) ======
  const formToEndpoint = {
    "form-manual-gq": "/gerar-portaria-gq",
    "form-excel-gq": "/gerar-portaria-gq-lote",

    "form-manual-remocao": "/gerar-portaria-movimentacao",
    "form-excel-remocao": "/gerar-portaria-movimentacao-lote",

    "form-manual-vacancia": "/gerar-portaria-vacancia",
    "form-excel-vacancia": "/gerar-portaria-vacancia-lote",

    "form-manual-gsiste": "/gerar-portaria-gsiste",
    "form-excel-gsiste": "/gerar-portaria-gsiste-lote",
  };

  function filenameFromDisposition(disposition) {
    // tenta extrair filename="..."
    if (!disposition) return null;
    const match = /filename\*?=(?:UTF-8''|")?([^\";]+)/i.exec(disposition);
    if (!match) return null;
    return decodeURIComponent(match[1].replace(/"/g, "").trim());
  }

  async function submitAndDownload(form) {
    const endpoint = formToEndpoint[form.id];
    if (!endpoint) {
      alert("Formulário sem rota configurada: " + form.id);
      return;
    }

    const formData = new FormData(form);

    const response = await fetch(endpoint, {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      // tenta ler mensagem do backend
      const text = await response.text().catch(() => "");
      alert("Erro ao gerar. " + (text ? ("\n\n" + text) : ""));
      return;
    }

    const blob = await response.blob();

    // nome do arquivo vindo do servidor (se existir)
    const cd = response.headers.get("content-disposition");
    const serverName = filenameFromDisposition(cd);

    // fallback por tipo
    const isZip = (response.headers.get("content-type") || "").includes("zip");
    const fallbackName = isZip ? "portarias.zip" : "portaria.docx";

    const filename = serverName || fallbackName;

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // intercepta submits dos forms do sistema
  Object.keys(formToEndpoint).forEach(formId => {
    const form = document.getElementById(formId);
    if (!form) return;

    form.addEventListener("submit", (e) => {
      e.preventDefault();
      submitAndDownload(form).catch(err => {
        console.error(err);
        alert("Erro inesperado ao gerar a portaria.");
      });
    });
  });
});
