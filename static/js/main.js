document.addEventListener("DOMContentLoaded", function () {
  // ---------- CONTROLE DAS ABAS ----------
  const tabs = document.querySelectorAll("[data-tab]");
  const sections = document.querySelectorAll(".tab-section");

  tabs.forEach(tab => {
    tab.addEventListener("click", function () {
      const target = this.getAttribute("data-tab");

      tabs.forEach(t => t.classList.remove("active-tab"));
      this.classList.add("active-tab");

      sections.forEach(section => {
        section.style.display = "none";
      });

      const activeSection = document.getElementById(target);
      if (activeSection) {
        activeSection.style.display = "block";
      }
    });
  });

  // Mostrar a primeira aba por padrão
  if (tabs.length > 0) {
    tabs[0].click();
  }

  // ---------- SUBMISSÃO DE FORMULÁRIOS ----------
  const forms = document.querySelectorAll("form");

  forms.forEach(form => {
    form.addEventListener("submit", async function (e) {
      e.preventDefault();

      const action = form.getAttribute("action");
      const method = form.getAttribute("method") || "POST";
      const formData = new FormData(form);

      try {
        const response = await fetch(action, {
          method: method,
          body: formData
        });

        if (!response.ok) {
          alert("Erro ao gerar a portaria.");
          return;
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement("a");
        a.href = url;
        a.download = "portaria.zip";
        document.body.appendChild(a);
        a.click();

        a.remove();
        window.URL.revokeObjectURL(url);
      } catch (err) {
        console.error(err);
        alert("Erro inesperado ao processar a solicitação.");
      }
    });
  });
});
