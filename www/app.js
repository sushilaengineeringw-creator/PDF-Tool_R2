// PDF.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
}

// ---------- Navigation ----------
function showSection(sectionId) {
  const sections = [
    'dashboard', 'viewer', 'img2pdf', 'merge', 'split',
    'compress', 'pdf2word', 'word2pdf', 'scanner'
  ];
  sections.forEach(sec => {
    const el = document.getElementById('sec-' + sec);
    if (el) el.classList.add('hidden');
  });
  const target = document.getElementById('sec-' + sectionId);
  if (target) target.classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---------- 1. PDF Viewer ----------
document.getElementById('viewer-input')?.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const container = document.getElementById('viewer-container');
  container.innerHTML = '<p class="text-slate-400 text-sm">Rendering pages...</p>';

  try {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    container.innerHTML = '';

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale: 1.3 });
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      canvas.height = viewport.height;
      canvas.width = viewport.width;
      await page.render({ canvasContext: ctx, viewport }).promise;
      container.appendChild(canvas);
    }
  } catch (err) {
    container.innerHTML = '<p class="text-rose-400 text-sm">Error loading PDF</p>';
    console.error(err);
  }
});

// ---------- 2. Photo to PDF ----------
async function processImg2Pdf() {
  const input = document.getElementById('img2pdf-input');
  if (!input.files || input.files.length === 0) {
    return alert('Please select at least one image.');
  }

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF();

  for (let i = 0; i < input.files.length; i++) {
    const dataUrl = await readFileAsDataURL(input.files[i]);
    if (i > 0) pdf.addPage();
    // Fit image nicely on A4
    const imgProps = pdf.getImageProperties(dataUrl);
    const pdfWidth = pdf.internal.pageSize.getWidth() - 20;
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    pdf.addImage(dataUrl, 'JPEG', 10, 10, pdfWidth, pdfHeight);
  }
  pdf.save('photos-converted.pdf');
}

// ---------- 3. Merge PDF ----------
async function processMergePdf() {
  const input = document.getElementById('merge-input');
  if (!input.files || input.files.length < 2) {
    return alert('Please select at least 2 PDF files.');
  }

  const mergedPdf = await PDFLib.PDFDocument.create();

  for (let i = 0; i < input.files.length; i++) {
    const bytes = await input.files[i].arrayBuffer();
    const pdf = await PDFLib.PDFDocument.load(bytes);
    const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    pages.forEach(p => mergedPdf.addPage(p));
  }

  const mergedBytes = await mergedPdf.save();
  downloadBlob(mergedBytes, 'merged-document.pdf', 'application/pdf');
}

// ---------- 4. Split PDF ----------
async function processSplitPdf() {
  const input = document.getElementById('split-input');
  if (!input.files || input.files.length === 0) {
    return alert('Please select a PDF file.');
  }

  const bytes = await input.files[0].arrayBuffer();
  const srcPdf = await PDFLib.PDFDocument.load(bytes);
  const total = srcPdf.getPageCount();

  for (let i = 0; i < total; i++) {
    const newPdf = await PDFLib.PDFDocument.create();
    const [page] = await newPdf.copyPages(srcPdf, [i]);
    newPdf.addPage(page);
    const out = await newPdf.save();
    downloadBlob(out, `page-${i + 1}.pdf`, 'application/pdf');
  }
  alert(`Split complete! ${total} files downloaded.`);
}

// ---------- 5. Compress PDF ----------
async function processCompressPdf() {
  const input = document.getElementById('compress-input');
  if (!input.files || input.files.length === 0) {
    return alert('Please select a PDF file.');
  }

  const bytes = await input.files[0].arrayBuffer();
  const pdfDoc = await PDFLib.PDFDocument.load(bytes);
  // Re-serialize with object streams → usually smaller
  const compressed = await pdfDoc.save({ useObjectStreams: true });
  downloadBlob(compressed, 'compressed.pdf', 'application/pdf');
}

// ---------- 6. PDF to Word ----------
async function processPdfToWord() {
  const input = document.getElementById('pdf2word-input');
  if (!input.files || input.files.length === 0) {
    return alert('Please select a PDF file.');
  }

  const bytes = await input.files[0].arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  let fullText = '';

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items.map(item => item.str).join(' ');
    fullText += `\n\n--- Page ${i} ---\n\n${pageText}`;
  }

  const doc = new docx.Document({
    sections: [{
      children: fullText.split('\n').map(line =>
        new docx.Paragraph({ children: [new docx.TextRun(line)] })
      )
    }]
  });

  const blob = await docx.Packer.toBlob(doc);
  downloadBlob(blob, 'converted.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

// ---------- 7. Text / Word to PDF ----------
async function processWordToPdf() {
  const text = document.getElementById('word2pdf-text').value;
  if (!text.trim()) return alert('Please enter some text.');

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF();
  const lines = pdf.splitTextToSize(text, 180);
  pdf.text(lines, 15, 20);
  pdf.save('document.pdf');
}

// ---------- 8. Camera Scan (web fallback) ----------
async function triggerCameraScan() {
  const preview = document.getElementById('scan-preview');
  preview.innerHTML = '<p class="text-slate-400">Opening camera...</p>';

  try {
    // Prefer Capacitor Camera if available
    if (window.Capacitor && window.Capacitor.Plugins && window.Capacitor.Plugins.Camera) {
      const photo = await window.Capacitor.Plugins.Camera.getPhoto({
        quality: 90,
        allowEditing: true,
        resultType: 'uri',
        source: 'CAMERA'
      });
      preview.innerHTML = `<img src="${photo.webPath}" class="max-w-full rounded-lg" />`;
      return;
    }

    // Browser fallback
    const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
    const video = document.createElement('video');
    video.srcObject = stream;
    video.autoplay = true;
    video.playsInline = true;
    video.className = 'max-w-full rounded-lg';
    preview.innerHTML = '';
    preview.appendChild(video);

    // Capture button
    const btn = document.createElement('button');
    btn.className = 'mt-3 btn bg-teal-600 hover:bg-teal-700';
    btn.innerHTML = '<i class="fa-solid fa-camera"></i> Take Photo';
    btn.onclick = () => {
      const canvas = document.createElement('canvas');
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      canvas.getContext('2d').drawImage(video, 0, 0);
      stream.getTracks().forEach(t => t.stop());
      const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
      preview.innerHTML = `<img src="${dataUrl}" class="max-w-full rounded-lg" />`;
      // Auto convert to PDF
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF();
      pdf.addImage(dataUrl, 'JPEG', 10, 10, 190, 0);
      pdf.save('scanned-document.pdf');
    };
    preview.appendChild(btn);
  } catch (err) {
    preview.innerHTML = '<p class="text-rose-400 text-sm">Camera not available or permission denied</p>';
    console.error(err);
  }
}

// ---------- Signature Pad ----------
const sigCanvas = document.getElementById('signature-pad');
let isDrawing = false;
if (sigCanvas) {
  const ctx = sigCanvas.getContext('2d');
  ctx.strokeStyle = '#ffffff';
  ctx.lineWidth = 2;
  ctx.lineCap = 'round';

  const start = (e) => {
    isDrawing = true;
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
  };
  const draw = (e) => {
    if (!isDrawing) return;
    const pos = getPos(e);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
  };
  const stop = () => { isDrawing = false; };

  function getPos(e) {
    const rect = sigCanvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return {
      x: (clientX - rect.left) * (sigCanvas.width / rect.width),
      y: (clientY - rect.top) * (sigCanvas.height / rect.height)
    };
  }

  sigCanvas.addEventListener('mousedown', start);
  sigCanvas.addEventListener('mousemove', draw);
  sigCanvas.addEventListener('mouseup', stop);
  sigCanvas.addEventListener('mouseleave', stop);
  sigCanvas.addEventListener('touchstart', (e) => { e.preventDefault(); start(e); });
  sigCanvas.addEventListener('touchmove', (e) => { e.preventDefault(); draw(e); });
  sigCanvas.addEventListener('touchend', stop);
}

function clearSignature() {
  if (!sigCanvas) return;
  const ctx = sigCanvas.getContext('2d');
  ctx.clearRect(0, 0, sigCanvas.width, sigCanvas.height);
}

function saveSignaturePdf() {
  if (!sigCanvas) return;
  const dataUrl = sigCanvas.toDataURL('image/png');
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF();
  pdf.setFontSize(14);
  pdf.text('Digital Signature', 15, 20);
  pdf.addImage(dataUrl, 'PNG', 15, 30, 120, 50);
  pdf.save('signature.pdf');
}

// ---------- Helpers ----------
function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function downloadBlob(data, fileName, mimeType) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 150);
}
