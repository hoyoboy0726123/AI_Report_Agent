// 產生完整「AI Report Agent 操作 SOP」Word 檔(總覽 + 6 頁 + 安裝附錄)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, LevelFormat, HeadingLevel, PageNumber, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, PageBreak,
} = require("docx");

const ASSET = path.join(__dirname, "sop_assets");
const CN = "Microsoft JhengHei";
const CONTENT_W = 9360;
const IMG_W = 600;

const DIMS = {
  "sop_overview.png": [1440, 900],
  "sop_w1.png": [1425, 1015], "sop_w2.png": [1440, 900], "sop_w3.png": [1440, 900],
  "sop_w4.png": [1440, 900], "sop_w5.png": [1440, 900],
  "sop_p_images.png": [1440, 900], "sop_p_transfer.png": [1425, 973],
  "sop_p_agent.png": [1440, 900], "sop_p_agent_demo.png": [1425, 921],
  "sop_p_automap.png": [1440, 900],
  "sop_p_aiengine.png": [1425, 1149], "sop_p_settings.png": [1440, 900],
};

function img(file) {
  const [w, h] = DIMS[file];
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 100, after: 80 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(ASSET, file)),
      transformation: { width: IMG_W, height: Math.round(IMG_W * h / w) },
      altText: { title: file, description: "screenshot", name: file } })],
  });
}
function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(t)] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] }); }
function p(t) { return new Paragraph({ spacing: { after: 120 }, children: parseRuns(t) }); }
function parseRuns(text) {
  const out = [];
  text.split(/(\*\*[^*]+\*\*)/).forEach((seg) => {
    if (!seg) return;
    if (seg.startsWith("**") && seg.endsWith("**")) out.push(new TextRun({ text: seg.slice(2, -2), bold: true }));
    else out.push(new TextRun(seg));
  });
  return out;
}
let stepCounter = 0;
function steps(items) {
  stepCounter += 1;
  const ref = "num" + stepCounter;
  return items.map((t) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60 }, children: parseRuns(t) }));
}
function tip(text) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { fill: "FFF7ED", type: ShadingType.CLEAR },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "F59E0B" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "F59E0B" },
        left: { style: BorderStyle.SINGLE, size: 18, color: "F59E0B" }, right: { style: BorderStyle.SINGLE, size: 4, color: "F59E0B" },
      },
      margins: { top: 80, bottom: 80, left: 160, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: "💡 提示  ", bold: true, color: "B45309" }), ...parseRuns(text)] })],
    })] })],
  });
}

const numberingConfigs = [];
for (let i = 1; i <= 20; i++) numberingConfigs.push({ reference: "num" + i,
  levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 600, hanging: 320 } } } }] });

const c = [];

// ===== 封面 =====
c.push(new Paragraph({ spacing: { before: 2600 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "AI Report Agent", bold: true, size: 60, color: "4F46E5" })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160 }, children: [new TextRun({ text: "操作 SOP・圖文操作手冊", bold: true, size: 38 })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240 }, children: [new TextRun({ text: "一份範本 + 一張 Excel,批次產出每列一份報告", size: 24, color: "64748B" })] }));
c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2200 }, children: [new TextRun({ text: "本手冊每一個功能頁皆以「實際截圖 + 紅框標註 + 編號步驟」呈現,照著點就會用。", size: 22, color: "64748B" })] }));

// ===== 介面總覽 =====
c.push(h1("介面總覽"));
c.push(p("開啟後是一個單頁工作台。**左側是 6 大功能分頁**,點選即切換;主要操作都在右側主畫面。整個程式只有一個後端視窗,前端在瀏覽器內執行。"));
c.push(img("sop_overview.png"));
c.push(...steps([
  "**報告精靈**:用範本 + Excel 批次產出每列一份報告(最常用)。",
  "**圖片對應**:把資料夾照片放進現成報告 / 換掉範本圖片標籤。",
  "**Excel 搬移**:把來源 Excel 欄位依對應搬到目標 Excel。",
  "**AI 助手**:用中文自然語言交辦,由 Agent 自動操作。",
  "**AI 引擎**:選擇 Gemini / 本機 Ollama 模型、審查與用量設定。",
  "**設定**:各功能共用的預設值(輸出資料夾、檔名、圖片寬度等)。",
]));
c.push(tip("左下角可切換**深色模式**。所有設定改了會即時儲存,各分頁共用同一份設定。"));

// ===== 一、報告精靈 =====
c.push(h1("一、報告精靈（批次產生報告）"));
c.push(p("用『一份範本 + 一張 Excel』批次產生「每一列一份」的報告,跟著畫面上的 5 個步驟走完即可。"));
c.push(p("適用情境:教育訓練證明、成績單、合約、通知書、證書……任何「同一種版型、只換資料」的大量文件。"));
c.push(tip("本範例以「教育訓練完成證明」示範:Word 範本含 {{姓名}}、{{部門}}、{{課程名稱}}、{{成績}}、{{完成日期}} 標籤,Excel 有 5 位學員,最後產出 5 份證明。"));

c.push(h2("步驟一、設定來源"));
c.push(img("sop_w1.png"));
c.push(...steps([
  "**報告類型(右上角)**:選 **Word / Excel / PPTX**。本例選 Word。",
  "**① 範本**:點「瀏覽」選含 {{欄位}} 標籤的 **.docx 範本**。",
  "**② 資料來源**:選資料 **Excel**,指定「**工作表**」與「**標題列**」(欄位名稱在第幾列)。",
  "**③ 輸出**:設定「**輸出資料夾**」與「**檔名規則**」(可用欄位名與 {index},例:{姓名}_訓練證明)。",
  "右側「**資料預覽**」確認欄位與列數正確後再往下。",
]));
c.push(h2("步驟二、對應標籤"));
c.push(img("sop_w2.png"));
c.push(...steps([
  "左側「**Excel 欄位**」:**點一個欄位**(變色=已選);已對應的欄位顯示**綠勾**。",
  "右側「**Word 範本內容**」:在段落點「**＋ 插入**」,把欄位以 {{欄位}} 寫進範本。",
  "(選用)按「**AI 建議對應**」自動把欄位對到範本。",
  "進階:在真正的 Word 視窗按 **Ctrl+Shift+M** 也能即時插入(需啟用熱鍵橋接)。",
]));
c.push(tip("若範本已寫好 {{標籤}}(像本範例),這步只是核對;欄位顯示綠勾即可直接下一步。"));
c.push(h2("步驟三、驗證"));
c.push(img("sop_w3.png"));
c.push(...steps([
  "進入此步會**自動檢查**範本標籤是否都能在 Excel 找到對應欄位。",
  "綠色「**全部對齊,可以產出!**」=可安全產出。",
  "「範本有、Excel 缺」=該標籤會**留空**;「Excel 有、範本沒用到」=不影響。",
  "調整後可按「**重新檢查**」。",
]));
c.push(h2("步驟四、產出"));
c.push(img("sop_w4.png"));
c.push(...steps([
  "確認範本/資料/輸出三路徑無誤,按「**開始產出**」。",
  "**進度條**顯示完成份數;右側「**即時產出清單**」逐一列出檔名。",
  "完成顯示「**產出 X/X 份**」,可按「**開啟輸出資料夾**」。",
  "若啟用 AI 審查,不合格的會自動移到 **Failed_Reports**。",
]));
c.push(h2("步驟五、完成 / 審查"));
c.push(img("sop_w5.png"));
c.push(...steps([
  "顯示整體流程完成:設定來源 → 對應標籤 → 驗證 → 產出 → 審查。",
  "可按「**開啟輸出資料夾**」檢視成果。",
  "要逐份自動評分,到「**AI 引擎**」開啟審查並選視覺模型。",
]));

// ===== 二、圖片對應 =====
c.push(h1("二、圖片對應（把照片放進報告）"));
c.push(p("把資料夾裡的照片塞進「**一份現成的報告檔**」。與報告精靈不同:精靈是批次產很多份,這裡是針對單一/制式報告配圖。"));
c.push(img("sop_p_images.png"));
c.push(p("**上方先選 3 種模式之一:**"));
c.push(...steps([
  "**制式範本(資料夾→各表)**:報告已含照片、位置固定(如測試報告);多個子資料夾(01、02…)各對一個照片表,依就近標籤換掉照片。",
  "**範本共用(配 Word)**:範本有 {{圖片欄}} 標籤,把固定一組圖依檔名換進去(例:公司 logo)。",
  "**視覺自動(Word/PPT・無標註)**:報告沒有任何標籤,AI 看版面自己決定每張圖貼哪。",
]));
c.push(p("**操作(以制式範本為例):**"));
c.push(...steps([
  "選「**制式範本**」(.docx / .pptx)。",
  "選「**照片根目錄**」(底下要有 01、02… 子資料夾,每個子資料夾 = 一個樣本)。",
  "選「**配對方式**」:auto(推薦,檔名語意優先)/ text(只看檔名)/ vlm(只看圖片內容)。",
  "先按「**① 先預覽對應(dry-run)**」確認配對無誤,再按「**② 正式產出**」。",
]));
c.push(tip("視覺自動 與 vlm 配對需要在「AI 引擎」先選好視覺(Reviewer)模型。"));

// ===== 三、Excel 搬移 =====
c.push(h1("三、Excel 搬移（欄位搬移彙整）"));
c.push(p("把**來源 Excel** 的欄位,依對應關係搬到**目標 Excel**(例:把多份名單彙整成一張總表)。"));
c.push(img("sop_p_transfer.png"));
c.push(...steps([
  "**來源**:選來源 Excel + 工作表 + 標題列。",
  "**目標**:選目標 Excel + 工作表(可新建)+ 標題列。",
  "**欄位對應**:按「**自動對應同名**」一鍵配好,或手動把每個來源欄指到目標欄(不想搬就留空)。",
  "**寫入模式**:**附加**(寫到既有資料下方)/ **覆寫**(清空標題列以下再寫)/ **全新**(新建目標檔)。",
  "按「**開始搬移**」。",
]));
c.push(tip("「自動對應同名」會把名稱一樣的欄位自動配好,最省事;不同名的再手動指定。"));

// ===== 四、AI 助手 =====
c.push(h1("四、AI 助手（一句話自動完成・最強用法）"));
c.push(p("這是本工具**最強大**的用法。前面幾個分頁要你「自己一步步點」;**AI 助手只要你打一句話,Agent 就會自己看、自己把整個流程做完。**"));
c.push(p("**重點:不必先進精靈對應標籤、也不必一步步設定。** 範本只要含有 {{標籤}},Agent 會自己讀目前設定、自己驗證對齊、自己批次產出、最後回報結果 —— 那些步驟你完全不用碰。"));
c.push(tip("使用前只需在「AI 引擎」選好 Planner 模型(本示範用雲端 Gemma-4 31B)。雲端模型速度快、工具呼叫穩定;本機 Ollama 也可,但需夠強的模型且 context 足夠(num_ctx=32768)。"));

c.push(h2("介面說明"));
c.push(img("sop_p_agent.png"));
c.push(...steps([
  "在輸入框打需求,按 **Ctrl+Enter** 或點**送出**。",
  "**對話區**會即時串流顯示 Agent 的思考、每一步**工具呼叫**與結果。",
  "要換個任務按「**新對話**」清空重來。",
]));

c.push(h2("實戰示範:一句話產出 5 份報告"));
c.push(p("以下是一次真實對話。使用者**只打了一句**「幫我用目前的設定,先驗證範本和資料有沒有對齊,沒問題就直接批次產出 Word 報告。」,接著完全由 Agent 自動完成:"));
c.push(img("sop_p_agent_demo.png"));
c.push(...steps([
  "**① 你只打這一句**:不用先對應標籤、不用進精靈設定。",
  "**② Agent 自己驗證**:自動呼叫 validate_template,確認範本標籤與 Excel 欄位完全對齊(passed: true)。",
  "**③ Agent 自己產出**:驗證通過後自動呼叫 generate_reports,一次批次產出 5 份(produced: 5)。",
  "**④ 回報完成**:用中文回報「範本與資料已成功對齊,報告已批次產出完畢」。",
]));
c.push(p("整個過程你沒有點任何按鈕、沒有手動對應任何欄位 —— Agent 看懂需求後,自己把該做的工具一個個叫出來執行完畢。"));

c.push(h2("進階示範:範本還沒標籤?AI 看懂兩個檔案自動補上"));
c.push(p("更強的是:**連範本都還沒對應標籤也行**。只要範本裡有「欄位名 / 標題」當線索(例:「姓名:」「部門:」),AI 會比對**範本內容**與 **Excel 欄位**,自動推斷該插哪些標籤。下圖是 AI 讀完一份**完全沒有 {{標籤}}** 的證明範本後的提案:"));
c.push(img("sop_p_automap.png"));
c.push(...steps([
  "**一句話**:「這份範本還沒有標籤,請看懂它和資料 Excel,幫我自動把對應標籤標上去。」",
  "AI 讀**範本內容 + Excel 欄位**,提出建議 —— 連**語意對應**都看得懂(例:範本的「測驗成績:」自動對到 Excel 的「成績」欄)。",
  "**人工檢查關卡(重要)**:AI 不會擅自改檔,而是先列出要插入的標籤、問你「是否全部套用」。你**看過確認無誤**再按「是,全部套用」(也可選「否,我想修改」)。",
  "套用後,AI 把標籤寫進**範本副本(檔名加 `_已標註`)**,你的**空白母版保持不動**;之後就能直接驗證 + 批次產出。",
]));
c.push(tip("**想提高成功率,範本至少要有「欄位名 / 標題」當線索**(如「姓名:」「部門:」)。完全空白、毫無文字提示的版面,AI 較難猜對位置 —— 建議先把欄位名打上,再交給 AI 自動補標籤。"));
c.push(tip("**第一次自動標註務必人工檢查**:確認 AI 插入的位置與欄位對應正確後再批次產出。標好的副本下次可直接重用,AI 不必再重看一次。"));

c.push(h2("你可以這樣『一句話』交辦"));
c.push(...steps([
  "「**現在的設定是什麼?**」→ 查目前範本 / 資料 / 輸出設定。",
  "「**驗證範本跟資料有沒有對齊**」→ 自動檢查標籤與欄位。",
  "「**用目前設定批次產生 Word 報告**」→ 直接產出。",
  "「**幫我把範本標籤對應好**」→ 自動把 Excel 欄位對應進範本,免進精靈手動點。",
  "「**改成產出 PPT 簡報**」/「**改用 Excel 範本**」→ 切換報告類型再產。",
  "「**把照片資料夾依檔名填進每一列的照片欄**」→ 自動配圖到資料來源。",
]));
c.push(tip("Agent 遇到不確定(例如標籤對不齊、缺檔)會主動發問請你確認,不會亂做;你回一句就繼續。"));

// ===== 五、AI 引擎 =====
c.push(h1("五、AI 引擎（模型與審查設定）"));
c.push(p("設定產報告 / 審查用的模型,以及 API 呼叫用量上限。"));
c.push(img("sop_p_aiengine.png"));
c.push(...steps([
  "**供應商與模型**:選 **Gemini**(需 API key)或本機 **Ollama**;按「**測試連線**」載入模型清單,再選 **Planner**(產報告/Agent)與 **Reviewer**(視覺審查)模型。",
  "**產出審查(選用)**:勾「**啟用審查**」後,產報告會逐份用視覺模型評分;可調**抽樣比例**與**評分標準(rubric)**。",
  "**呼叫預算**:設定 Planner / Reviewer 呼叫上限與**視覺配速 RPM**,避免 API 用量失控。",
]));
c.push(tip("Gemini 金鑰放在專案根目錄 .env 的 GEMINI_API_KEY(參考 .env.example);Ollama 免金鑰,確認 Endpoint(預設 http://localhost:11434)即可。"));

// ===== 六、設定 =====
c.push(h1("六、設定（共用預設值）"));
c.push(p("各功能共用的預設值,**改了會立即儲存**。"));
c.push(img("sop_p_settings.png"));
c.push(...steps([
  "**輸出預設**:預設輸出資料夾、Word / Excel 報告檔名規則。",
  "**圖片嵌入**:Word 圖片寬度(mm)、Excel 圖片寬度(px)。",
  "**熱鍵橋接(進階)**:設定 Ctrl+Shift+M 在真正 Office 視窗插入標籤的目標(需另啟動橋接程式)。",
  "**關於**:版本與專案資訊。",
]));

// ===== 附錄:安裝與啟動 =====
c.push(h1("附錄、安裝與啟動"));
c.push(h2("A. 安裝後端依賴(擇一)"));
c.push(p("**方式一:uv(建議,免煩惱 Python 版本)** — uv 會依 .python-version 自動下載對的 Python 並安裝鎖定版依賴。"));
c.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "uv sync", font: "Consolas", size: 20 })] }));
c.push(p("**方式二:傳統 pip** — 需本機已有 Python 3.11+(建議 3.13)。"));
c.push(new Paragraph({ children: [new TextRun({ text: "py -3.13 -m venv .venv", font: "Consolas", size: 20 })] }));
c.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: ".venv\\Scripts\\python -m pip install -r backend\\requirements.txt", font: "Consolas", size: 20 })] }));
c.push(h2("B. 啟動"));
c.push(...steps([
  "雙擊 **launch.bat** → 自動啟動後端(http://127.0.0.1:8756)並開啟瀏覽器。",
  "整個 app **只有一個程序**:後端同時供應 API 與已 build 的前端網頁;那個黑色命令視窗關掉就等於關閉服務。",
]));
c.push(h2("C. 看錯誤訊息"));
c.push(...steps([
  "**後端**錯誤(Python / API / 產出):看那個黑色命令視窗。",
  "**前端**錯誤(畫面 / 按鈕):在瀏覽器按 **F12** → Console(JS 錯誤)與 Network(API 失敗)。",
]));
c.push(h2("D. 設定 API key"));
c.push(p("在專案根目錄建立 **.env**(參考 .env.example),填入 GEMINI_API_KEY;使用本機 Ollama 則免金鑰。"));

const doc = new Document({
  styles: {
    default: { document: { run: { font: CN, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: CN, color: "1E293B" }, paragraph: { spacing: { before: 120, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: CN, color: "4F46E5" }, paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: numberingConfigs },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "AI Report Agent 操作 SOP", size: 16, color: "94A3B8" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: "第 ", size: 16, color: "94A3B8" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "94A3B8" }), new TextRun({ text: " 頁", size: 16, color: "94A3B8" })] })] }) },
    children: c,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(__dirname, "AI_Report_Agent_操作SOP.docx");
  fs.writeFileSync(out, buf);
  console.log("WROTE", out, buf.length, "bytes");
});
