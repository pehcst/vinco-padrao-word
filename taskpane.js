// Painel "Padrão Vinco": cola, gera o .docx no próprio navegador e abre no Word.
const $ = (id) => document.getElementById(id);
let noWord = false;

function status(msg, tipo) { const s = $("status"); s.textContent = msg; s.className = tipo || ""; }
function avisos(lista) {
  const ul = $("avisos"); ul.innerHTML = "";
  for (const a of lista.slice(0, 6)) { const li = document.createElement("li"); li.textContent = a; ul.appendChild(li); }
}

// Avisos técnicos do gerador → linguagem simples (os mesmos lembretes do botão do Word)
function simples(av, texto) {
  const out = [];
  if (/\[[^\]\n]{2,60}\]/.test(texto.replace(/\[[^\]]*\]\([^)]*\)/g, ""))) out.push("Preencha os campos que ficaram [entre colchetes].");
  if (av.some((a) => /^Título 1/.test(a))) out.push("Algum título de capítulo está longo. O ideal é até 6 palavras.");
  if (av.some((a) => /^Título [23]/.test(a))) out.push("Algum subtítulo está longo. O ideal é até 8 palavras.");
  if (av.some((a) => /^Tabela sem identificação/.test(a))) out.push("Tem tabela sem nome. Escreva o nome dela na linha de cima.");
  if (av.some((a) => /^(Parágrafo|Frase)/.test(a))) out.push("Tem parágrafo ou frase longa. Vale quebrar em partes menores.");
  if (av.some((a) => /^Imagem/.test(a))) out.push("Imagens não entram sozinhas: insira pelo Word (Inserir → Imagens).");
  return out;
}

async function abrirNoWord(base64) {
  await Word.run(async (ctx) => { ctx.application.createDocument(base64).open(); await ctx.sync(); });
}

// Documento aberto vazio? Preenche ele mesmo (sem as mensagens do Word). Com texto, cria um novo para não apagar nada.
async function documentoVazio() {
  return Word.run(async (ctx) => { const b = ctx.document.body; b.load("text"); const t = b.getRange().inlinePictures; t.load("items"); const tb = b.tables; tb.load("items"); await ctx.sync(); return !b.text.trim() && !t.items.length && !tb.items.length; });
}
async function preencherAtual(base64) {
  await Word.run(async (ctx) => { ctx.document.body.insertFileFromBase64(base64, "Replace"); await ctx.sync(); });
}

function oferecerDownload(base64) {
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
  const a = $("baixar"); a.href = url; a.download = "Documento Vinco.docx"; a.hidden = false;
}

async function criar() {
  const texto = $("texto").value;
  $("baixar").hidden = true; avisos([]);
  if (!texto.trim()) { status("Você ainda não colou nada. Copie a resposta no Copilot e cole na caixa acima.", "erro"); $("texto").focus(); return; }
  $("criar").disabled = true; status("Criando o documento…");
  try {
    const r = await VincoCore.gerarDocumento(texto);
    const lembretes = simples(r.avisos, texto);
    if (!r.temCabecalho) lembretes.unshift("Esse texto não veio do agente Documentos Vinco, então saiu um documento simples, sem capa.");
    if (noWord) {
      let feito = false;
      if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
        try { if (await documentoVazio()) { await preencherAtual(r.base64); feito = true; status("Pronto! O documento está no padrão Vinco.", "ok"); } }
        catch (e) { console.error("preencher atual", e); }
      }
      if (feito) {}
      else try { await abrirNoWord(r.base64); status("Pronto! O documento abriu numa nova janela ou aba.", "ok"); }
      catch (e) { oferecerDownload(r.base64); status("O documento está pronto. Clique em Baixar o documento.", "ok"); console.error(e); }
    } else { oferecerDownload(r.base64); status("O documento está pronto. Clique em Baixar o documento.", "ok"); }
    avisos(lembretes);
  } catch (e) {
    console.error(e);
    status(e.message === "vazio" ? "Não encontrei texto para formatar. Copie de novo no Copilot." : "Não consegui criar o documento. Copie de novo pelo botão Copiar do Copilot e tente outra vez.", "erro");
  } finally { $("criar").disabled = false; }
}

$("criar").addEventListener("click", criar);
if (window.Office) Office.onReady((info) => { noWord = info.host === Office.HostType.Word && Office.context.requirements.isSetSupported("WordApi", "1.3"); });
