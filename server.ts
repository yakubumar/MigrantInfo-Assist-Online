import "dotenv/config";
import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import { 
  Document, 
  Packer, 
  Paragraph, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  TextRun, 
  BorderStyle,
  PageOrientation
} from "docx";

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || "0.0.0.0";

app.use(express.json({ limit: "15mb" }));
app.use(express.urlencoded({ extended: true, limit: "15mb" }));

// Gemini клиенти биринчи сўровда яратилади (калит бўлмаса дастур ўчиб қолмайди)
let aiClient: GoogleGenAI | null = null;
function getAI(): GoogleGenAI {
  const key = (process.env.GEMINI_API_KEY || "").trim();
  if (!key || key === "MY_GEMINI_API_KEY") {
    throw new Error("GEMINI_API_KEY топилмади. Дастур папкасидаги .env файлига Gemini API калитини ёзинг (README.md га қаранг).");
  }
  if (!aiClient) {
    aiClient = new GoogleGenAI({
      apiKey: key,
      httpOptions: { headers: { 'User-Agent': 'aistudio-build' } }
    });
  }
  return aiClient;
}

// Helper to remove literal "\n\n" string artifacts and clean text
function cleanLiteralNewlines(text: string): string {
  if (!text) return "";
  return text
    .replace(/\\n/g, '\n')
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Server-side helper to fetch and extract article text from a URL with robust charset detection (windows-1251, utf-8, etc.)
async function extractArticleFromUrl(url: string): Promise<string> {
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 12000);
    const res = await fetch(url, {
      signal: controller.signal,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,uz;q=0.7'
      }
    });
    clearTimeout(timeout);
    if (!res.ok) return "";

    const arrayBuffer = await res.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Detect charset from headers or HTML meta tag
    const contentType = res.headers.get('content-type') || '';
    let charset = 'utf-8';
    const headerMatch = contentType.match(/charset=([a-zA-Z0-9_-]+)/i);
    if (headerMatch) {
      charset = headerMatch[1].toLowerCase();
    } else {
      // Look at first 2048 bytes as ascii for meta charset or http-equiv
      const initialChunk = buffer.subarray(0, 2048).toString('ascii');
      const metaMatch = initialChunk.match(/charset=["']?([a-zA-Z0-9_-]+)/i);
      if (metaMatch) {
        charset = metaMatch[1].toLowerCase();
      }
    }

    const decodeText = (cs: string): string => {
      try {
        const decoder = new TextDecoder(cs);
        return decoder.decode(buffer);
      } catch {
        return buffer.toString('utf-8');
      }
    };

    let html = decodeText(charset);

    // If decoded HTML has high ratio of replacement characters (e.g. windows-1251 misidentified as utf-8)
    const replacementCount = (html.match(/\uFFFD/g) || []).length;
    if (replacementCount > 10 && charset !== 'windows-1251') {
      const fallbackHtml = decodeText('windows-1251');
      const fallbackReplacements = (fallbackHtml.match(/\uFFFD/g) || []).length;
      if (fallbackReplacements < replacementCount) {
        html = fallbackHtml;
      }
    }

    // Strip non-content blocks (scripts, styles, comments, nav, footer, header, aside)
    const textContent = html
      .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
      .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, ' ')
      .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, ' ')
      .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, ' ')
      .replace(/<aside\b[^<]*(?:(?!<\/aside>)<[^<]*)*<\/aside>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/\s+/g, ' ')
      .trim();

    return textContent.slice(0, 25000);
  } catch (err) {
    console.warn("Direct URL extraction failed:", err);
    return "";
  }
}

// Server-side helper to format Gemini errors into clear Uzbek messages
function formatGeminiError(err: any): string {
  const msg = err?.message || String(err || "");
  if (
    msg.includes("429") ||
    msg.includes("RESOURCE_EXHAUSTED") ||
    msg.includes("quota") ||
    msg.includes("rate-limits")
  ) {
    return "Gemini API сўровлар лимитига (квота/429) дуч келинди. Бу вақтинчалик чеклав (1 дақиқалик сўровлар миқдори). Илтимос, 15-20 сония кутиб, тугмани қайта босинг.";
  }
  if (
    msg.includes("503") ||
    msg.includes("UNAVAILABLE") ||
    msg.includes("high demand")
  ) {
    return "Сунъий интеллект сервери ҳозирда юқори юклама остида (503). Илтимос, 5-10 сониядан сўнг қайта уриниб кўринг.";
  }
  return msg;
}

// Resilient Gemini generateContent helper with automatic model fallback, retry and rate-limit backoff
async function generateWithFallback(params: { contents: any; config?: any }) {
  // Valid, supported models ordered by responsiveness and availability: gemini-3.1-flash-lite, gemini-flash-latest, gemini-3.8-flash
  const candidateModels = ["gemini-3.1-flash-lite", "gemini-flash-latest", "gemini-3.8-flash"];
  let lastError: any = null;

  // Try up to 2 passes across candidate models to withstand brief infrastructure spikes
  for (let pass = 1; pass <= 2; pass++) {
    for (const modelName of candidateModels) {
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          const callPromise = getAI().models.generateContent({
            model: modelName,
            contents: params.contents,
            config: params.config
          });
          const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error(`Model ${modelName} call timed out after 18s`)), 18000)
          );
          const resp: any = await Promise.race([callPromise, timeoutPromise]);
          if (resp && typeof resp.text === "string" && resp.text.trim()) {
            return resp;
          }
        } catch (err: any) {
          lastError = err;
          const errMsg = String(err?.message || "");
          const isRateLimit = errMsg.includes("429") || errMsg.includes("RESOURCE_EXHAUSTED") || errMsg.includes("quota");
          const isUnavailable = errMsg.includes("503") || errMsg.includes("UNAVAILABLE") || errMsg.includes("high demand");

          // On attempt 1, if it's transient high demand (503) or rate limit (429), pause and retry the same model
          if (attempt === 1 && (isRateLimit || isUnavailable)) {
            await new Promise(res => setTimeout(res, 1200));
            continue;
          }

          // If attempt 2 or other error, short pause before moving to the next candidate model
          if (isRateLimit || isUnavailable) {
            await new Promise(res => setTimeout(res, 500));
          }
          break; // move to next model
        }
      }
    }

    if (pass < 2) {
      // Pause 1.5s before second pass across models if all experienced a transient spike
      await new Promise(res => setTimeout(res, 1500));
    }
  }

  throw new Error(formatGeminiError(lastError));
}

// API Health Check
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", hasApiKey: Boolean(process.env.GEMINI_API_KEY) });
});

// Process and Rewrite Article + Generate Headlines in ONE efficient single turn
app.post("/api/process-article", async (req, res) => {
  try {
    const { inputMode, url, inputText, stylePrompt } = req.body;

    const customInstruction = stylePrompt && stylePrompt.trim()
      ? `\n\nМУҲИМ МАХСУС УСЛУБ ВА КЎРСАТМАЛАР (МУАЛЛИФДАН):\n"""\n${stylePrompt.trim()}\n"""\nУшбу кўрсатма ва услубий талабларга ҚАТЪИЙ АМАЛ ҚИЛ! Текстни айнан шу кўрсатилган услуб ва оҳангда қайта ишла.`
      : "";

    let articleSourceText = "";

    if (inputMode === 'url') {
      if (!url || !url.trim().startsWith("http")) {
        return res.status(400).json({ error: "Илтимос, тўғри веб-ҳавола (URL) киритинг." });
      }

      const extractedContent = await extractArticleFromUrl(url.trim());
      
      if (!extractedContent || extractedContent.length < 60) {
        return res.status(400).json({
          error: "Мазкур ҳаволадан мақола матнини автоматик ўқиб бўлмади (сайт ҳимояси ёки саҳифа ёпиқ). Илтимос, сайтга кириб мақола матнини нусхаланг ва «Матн киритиш» бўлимига қўйинг."
        });
      }

      articleSourceText = extractedContent;
    } else {
      if (!inputText || !inputText.trim()) {
        return res.status(400).json({ error: "Илтимос, қайта ишлаш учун мақола матнини киритинг." });
      }
      articleSourceText = inputText.trim();
    }

    // Unified single-call prompt: rewrites article AND generates 10 headlines together
    // This halves API requests (prevents 429 quota exhaustion) and speeds up response time to ~2-3s!
    const prompt = `Прочитай следующий текст статьи:
---
${articleSourceText}
---
Выполни две задачи:
1. ПЕРЕПИШИ СТАТЬЮ: Перепиши ее как простую и понятную статью для мигрантов. Используй ясный и доступный русский язык. Предложения должны быть короткими и понятными, но не детскими. 
ПРЯМОЕ ТРЕБОВАНИЕ: обязательно сохрани и включи в текст абсолютно ВСЕ цифры, даты, суммы и статистические данные из оригинального источника. 
НЕ объясняй общеизвестные термины, такие как "выдворение", "депортация", "QR-код" или "запрет на въезд". 
Текст должен быть цельным, без подзаголовков и без списков. 
КРИТИЧЕСКОЕ ТРЕБОВАНИЕ: Ни в коем случае НЕ начинай статью со стандартных клише или вводных фраз о дате/времени (например, категорически типа "По данным на [дата/время]...", "По состоянию на...", "На момент..." или упоминаний текущего дня). Начинай повествование сразу с сути новости или главного события.
ЗАПРЕТ НА ВРЕМЯ И ДАТУ ПУБЛИКАЦИИ: Ни в коем случае не указывай дату и точное время публикации статьи, беседы, интервью или выхода новости (например, НЕ пиши: "которая вышла 14 июня 2026 года в 17:08"). Вместо этого пиши просто и без указания времени/даты публикации, например: "Об этом рассказал Борис Титов в беседе с ТАСС".${customInstruction}

2. ПРИДУМАЙ 10 ЗАГОЛОВКОВ: Придумай ровно 10 привлекательных, емких и понятных заголовков для этой статьи на простом русском языке.${stylePrompt && stylePrompt.trim() ? ` Учитывай заданный стиль: ${stylePrompt.trim()}` : ''}

Верни результат строго в формате JSON со следующими полями:
- "rewrittenText": полный переписанный текст статьи (строка)
- "headlines": массив из 10 привлекательных заголовков (массив строк)`;

    const response = await generateWithFallback({
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            rewrittenText: { type: Type.STRING, description: "Переписанный текст статьи" },
            headlines: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
              description: "10 емких заголовков"
            }
          },
          required: ["rewrittenText", "headlines"]
        }
      }
    });

    let rewrittenText = "";
    let headlines: string[] = [];

    try {
      const data = JSON.parse(response.text || "{}");
      rewrittenText = cleanLiteralNewlines(data.rewrittenText || "");
      headlines = Array.isArray(data.headlines) ? data.headlines.map(cleanLiteralNewlines) : [];
    } catch {
      rewrittenText = cleanLiteralNewlines(response.text || "");
      headlines = [rewrittenText.slice(0, 80)];
    }

    if (!rewrittenText) {
      return res.status(500).json({ error: "Мақола матнини қайта ишлашда хатолик юз берди." });
    }

    if (!headlines || headlines.length === 0) {
      headlines = ["Янги хабар ва тафсилотлар"];
    }

    res.json({
      rewrittenText,
      headlines
    });
  } catch (err: any) {
    console.error("Error in /api/process-article:", err);
    res.status(500).json({ error: formatGeminiError(err) });
  }
});

// Regenerate 10 new, fresh headline options
app.post("/api/regenerate-headlines", async (req, res) => {
  try {
    const { rewrittenText, stylePrompt, currentHeadlines } = req.body;

    if (!rewrittenText || !rewrittenText.trim()) {
      return res.status(400).json({ error: "Мақола матни кўрсатилмаган." });
    }

    const previousList = Array.isArray(currentHeadlines) && currentHeadlines.length > 0
      ? `\nАввалги сарлавҳалар рўйхати (ушбу вариантлардан фарқ қилувчи, янгича ва ўзига хос ракурсдаги янги 10 та сарлавҳа таклиф қил):\n${currentHeadlines.map((h: string, i: number) => `- ${h}`).join('\n')}\n`
      : "";

    const customInstruction = stylePrompt && stylePrompt.trim()
      ? `\nМуаллиф услуби бўйича талаблар: "${stylePrompt.trim()}".`
      : "";

    const prompt = `Для следующей статьи придумай 10 НОВЫХ, привлекательных, емких, понятных и разнообразных заголовков на простом русском языке для мигрантов.${previousList}${customInstruction}

Текст статьи:
---
${rewrittenText.trim()}
---

Заголовки должны быть цепляющими, информативными, отражать главную суть статьи и не содержать кликбейта.
Верни результат строго в формате JSON со следующим полем:
- "headlines": массив из 10 новых заголовков (строк)`;

    const response = await generateWithFallback({
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            headlines: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
              description: "10 новых привлекательных заголовков"
            }
          },
          required: ["headlines"]
        }
      }
    });

    let headlines: string[] = [];
    try {
      const data = JSON.parse(response.text || "{}");
      headlines = Array.isArray(data.headlines) ? data.headlines.map(cleanLiteralNewlines) : [];
    } catch {
      headlines = [];
    }

    if (!headlines || headlines.length === 0) {
      return res.status(500).json({ error: "Янги сарлавҳаларни шакллантиришда хатолик юз берди." });
    }

    res.json({ headlines });
  } catch (err: any) {
    console.error("Error in /api/regenerate-headlines:", err);
    res.status(500).json({ error: formatGeminiError(err) });
  }
});

// Translate to Uzbek Cyrillic
app.post("/api/translate-uzbek", async (req, res) => {
  try {
    const { fullText, stylePrompt } = req.body;
    const customInstruction = stylePrompt && stylePrompt.trim()
      ? `\n6. МУҲИМ УСЛУБИЙ КЎРСАТМА: Таржимада мақоланинг ушбу услубига риоя қил: "${stylePrompt.trim()}".`
      : '';

    const response = await generateWithFallback({
      contents: `Переведи следующий текст строго по источнику на узбекский язык, используя КИРИЛЛИЦУ (кирилл алифбоси).
      
ТАБЛИЦА ПРАВИЛ ДЛЯ ПЕРЕВОДА:
1. Переводи СТРОГО по предоставленному тексту-источнику. Не добавляй от себя никаких новых фактов, разъяснений, контекста или собственных дополнений. Перевод должен точно и полно передавать смысл и структуру оригинала.
2. Не сокращай и не урезай текст — он должен быть полным и точным эквивалентом оригинала на узбекском языке.
3. ОБЯЗАТЕЛЬНО сохрани все цифры, даты, имена собственные и статистические данные в точности как в оригинале.
4. ВАЖНО: В кириллице ставь дефис только после цифр года и даты (например: 2025-йил, 25-май).
5. Не добавляй никаких вступительных или заключительных слов (например, "Вот перевод:", "Tarjima:"), комментариев переводчика или примечаний. Верни только чистый переведенный текст.${customInstruction}

Текст-источник для перевода:
${fullText}`,
    });

    const cyrillic = cleanLiteralNewlines(response.text?.trim() || "");
    res.json({ uzbekCyrillic: cyrillic });
  } catch (err: any) {
    console.error("Error in /api/translate-uzbek:", err);
    res.status(500).json({ error: formatGeminiError(err) });
  }
});

// Synchronize Russian text from edited Uzbek
app.post("/api/sync-russian", async (req, res) => {
  try {
    const { uzbekCyrillic, headline, russianText } = req.body;

    const response = await generateWithFallback({
      contents: `У тебя есть отредактированный текст на узбекском языке (кириллица) и его исходная русская версия.
Твоя задача — обновить, исправить и синхронизировать русскую версию (и заголовок, и текст) так, чтобы она ТОЧНО соответствовала узбекскому тексту по смыслу, фактам, структуре, цифрам и любым внесенным изменениям.

Отредактированный узбекский текст:
${uzbekCyrillic}

Исходная русская версия:
Заголовок: ${headline}
Текст:
${russianText}

Обнови русскую версию так, чтобы она стала точным и качественным эквивалентом узбекского текста на русском языке. Сохраняй стиль изложения. Возвращай результат СТРОГО в формате JSON с полями "headline" и "text".`,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            headline: { type: Type.STRING },
            text: { type: Type.STRING }
          },
          required: ["headline", "text"]
        }
      }
    });

    const data = JSON.parse(response.text || "{}");
    res.json({
      headline: cleanLiteralNewlines(data.headline?.trim() || ""),
      text: cleanLiteralNewlines(data.text?.trim() || "")
    });
  } catch (err: any) {
    console.error("Error in /api/sync-russian:", err);
    res.status(500).json({ error: formatGeminiError(err) });
  }
});

// Safe ASCII transliteration for RFC 6266 Content-Disposition fallback
function toSafeAsciiFilename(str: string): string {
  const map: Record<string, string> = {
    'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'yo','ж':'zh','з':'z',
    'и':'i','й':'y','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r',
    'с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'shch',
    'ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya','ў':'o','ғ':'g','қ':'q','ҳ':'h'
  };
  const lower = str.toLowerCase();
  let res = '';
  for (const ch of lower) {
    if (map[ch] !== undefined) {
      res += map[ch];
    } else if (/[a-z0-9_-]/.test(ch)) {
      res += ch;
    } else if (/\s/.test(ch)) {
      res += '_';
    }
  }
  const clean = res.replace(/_+/g, '_').replace(/^_+|_+$/g, '').slice(0, 30);
  return clean || 'maqola';
}

function cyrillicToLatinUz(text: string): string {
  if (!text) return '';
  const map: Record<string, string> = {
    'А': 'A', 'а': 'a', 'Б': 'B', 'б': 'b', 'В': 'V', 'в': 'v', 'Г': 'G', 'г': 'g',
    'Д': 'D', 'д': 'd', 'Е': 'E', 'е': 'e', 'Ё': 'Yo', 'ё': 'yo', 'Ж': 'J', 'ж': 'j',
    'З': 'Z', 'з': 'z', 'И': 'I', 'и': 'i', 'Й': 'Y', 'й': 'y', 'К': 'K', 'к': 'k',
    'Л': 'L', 'л': 'l', 'М': 'M', 'м': 'm', 'Н': 'N', 'н': 'n', 'О': 'O', 'о': 'o',
    'П': 'P', 'п': 'p', 'Р': 'R', 'р': 'r', 'С': 'S', 'с': 's', 'Т': 'T', 'т': 't',
    'У': 'U', 'у': 'u', 'Ф': 'F', 'ф': 'f', 'Х': 'X', 'х': 'x', 'Ц': 'Ts', 'ц': 'ts',
    'Ч': 'Ç', 'ч': 'ç', 'Ш': 'Ş', 'ш': 'ş', 'Щ': 'Ş', 'щ': 'ş', 'Ъ': "'", 'ъ': "'",
    'Ь': '', 'ь': '', 'Э': 'E', 'э': 'e', 'Ю': 'Yu', 'ю': 'yu', 'Я': 'Ya', 'я': 'ya',
    'Ў': 'Ö', 'ў': 'ö', 'Қ': 'Q', 'қ': 'q', 'Ғ': 'Ğ', 'ғ': 'ğ', 'Ҳ': 'H', 'ҳ': 'h',
  };

  let transliterated = text.split('').map((char, i, arr) => {
    const prev = arr[i - 1];
    const isVowel = prev && 'АаЕеЁёИиОоУуЭэЮюЯяЎўÖö'.includes(prev);
    const isStart = !prev || /[\s\p{P}]/u.test(prev);
    if ((char === 'Е' || char === 'е') && (isStart || isVowel)) {
      return char === 'Е' ? 'Ye' : 'ye';
    }
    if ((char === 'Ц' || char === 'ц') && isStart) {
      return char === 'Ц' ? 'S' : 's';
    }
    return map[char] !== undefined ? map[char] : char;
  }).join('');

  return transliterated
    .replace(/[OО][ʻ‘’’'`ʼ]/g, 'Ö')
    .replace(/[oо][ʻ‘’’'`ʼ]/g, 'ö')
    .replace(/[GГ][ʻ‘’’'`ʼ]/g, 'Ğ')
    .replace(/[gг][ʻ‘’’'`ʼ]/g, 'ğ')
    .replace(/SH/g, 'Ş')
    .replace(/Sh/g, 'Ş')
    .replace(/sh/g, 'ş')
    .replace(/CH/g, 'Ç')
    .replace(/Ch/g, 'Ç')
    .replace(/ch/g, 'ç');
}

// Word (.docx): фақат мақола — битта жадвал, икки устун (русча | ўзбекча).
// Ҳар бир устунда: қалин сарлавҳа + мақола матни. Устун номлари ва манба қатори йўқ.
function createWordDocument({
  headline = "",
  russianText = "",
  uzbekCyrillic = "",
  uzbekLatin = "",
  script = "cyrillic"
}: {
  headline?: string;
  russianText?: string;
  uzbekCyrillic?: string;
  uzbekLatin?: string;
  script?: "cyrillic" | "latin" | "both";
  sourceUrl?: string;
}) {
  const cleanLines = (text: string) => {
    if (!text) return [];
    return text
      .split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0 && !/^(источник|manba|манба):/i.test(line));
  };

  const ruLines = cleanLines(russianText);
  let uzCyrLines = cleanLines(uzbekCyrillic);
  let uzLatLines = cleanLines(uzbekLatin);

  if (uzCyrLines.length === 0 && uzLatLines.length > 0) {
    uzCyrLines = uzLatLines;
  } else if (uzLatLines.length === 0 && uzCyrLines.length > 0) {
    uzLatLines = uzCyrLines.map(l => cyrillicToLatinUz(l));
  }

  const hl = headline.trim();
  const ruHeadline = hl || ruLines[0] || "";
  const ruBody = ruLines[0] === ruHeadline ? ruLines.slice(1) : ruLines;
  // Ўзбекча матннинг биринчи қатори — таржима қилинган сарлавҳа
  const cyr = { headline: uzCyrLines[0] || "", body: uzCyrLines.slice(1) };
  const lat = { headline: uzLatLines[0] || "", body: uzLatLines.slice(1) };

  const columns: { headline: string; body: string[] }[] = script === "both"
    ? [{ headline: ruHeadline, body: ruBody }, cyr, lat]
    : [{ headline: ruHeadline, body: ruBody }, script === "latin" ? lat : cyr];

  const colWidth = Math.floor(100 / columns.length);
  const cellMargins = { top: 120, bottom: 120, left: 140, right: 140 };
  const line = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const cellBorders = { top: line, bottom: line, left: line, right: line };

  const cells = columns.map(col => {
    const paras: Paragraph[] = [];
    if (col.headline) {
      paras.push(new Paragraph({
        children: [new TextRun({ text: col.headline, bold: true, size: 28, font: "Calibri" })],
        spacing: { after: 280, line: 276 }
      }));
    }
    for (const t of col.body) {
      paras.push(new Paragraph({
        children: [new TextRun({ text: t, size: 28, font: "Calibri" })],
        spacing: { after: 200, line: 276 }
      }));
    }
    if (paras.length === 0) paras.push(new Paragraph({ children: [] }));
    return new TableCell({
      width: { size: colWidth, type: WidthType.PERCENTAGE },
      borders: cellBorders,
      margins: cellMargins,
      children: paras
    });
  });

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { orientation: columns.length === 3 ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
          margin: { top: 720, right: 720, bottom: 720, left: 720 }
        }
      },
      children: [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [new TableRow({ children: cells })]
        })
      ]
    }]
  });
}

// In-memory temporary cache for prepared Word files
interface PreparedFile {
  buffer: Buffer;
  filename: string;
  asciiName: string;
  timestamp: number;
}
const preparedFiles = new Map<string, PreparedFile>();

// Clean up expired files (> 20 mins)
setInterval(() => {
  const now = Date.now();
  for (const [id, item] of preparedFiles.entries()) {
    if (now - item.timestamp > 20 * 60 * 1000) {
      preparedFiles.delete(id);
    }
  }
}, 60 * 1000);

// Endpoint 1: Prepare Word docx and return download ID and direct URL
app.post("/api/prepare-docx", async (req, res) => {
  try {
    const {
      headline = "",
      russianText = "",
      uzbekCyrillic = "",
      uzbekLatin = "",
      script = "cyrillic",
      sourceUrl = ""
    } = req.body;

    const doc = createWordDocument({
      headline,
      russianText,
      uzbekCyrillic,
      uzbekLatin,
      script,
      sourceUrl
    });

    const buffer = await Packer.toBuffer(doc);
    const cleanPrefix = (headline || "Maqola")
      .split(/\s+/)
      .slice(0, 4)
      .join(" ")
      .replace(/[^\p{L}\p{N}\s_-]/gu, "")
      .trim() || "Maqola";
    const filename = `${cleanPrefix}.docx`;
    const asciiName = `${toSafeAsciiFilename(cleanPrefix)}.docx`;

    const id = Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
    preparedFiles.set(id, {
      buffer,
      filename,
      asciiName,
      timestamp: Date.now()
    });

    res.json({
      success: true,
      id,
      filename,
      downloadUrl: `/api/download-docx-file/${id}`
    });
  } catch (err: any) {
    console.error("Error in /api/prepare-docx:", err);
    res.status(500).json({ error: "Word ҳужжатини тайёрлашда хатолик: " + err.message });
  }
});

// Endpoint 2: Direct HTTP download from prepared ID
app.get("/api/download-docx-file/:id", (req, res) => {
  const item = preparedFiles.get(req.params.id);
  if (!item) {
    return res.status(404).send("Файл муддати тугаган ёки топилмади. Саҳифадан файлни қайта юкланг.");
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  // RFC 6266 and RFC 5987 compliant Content-Disposition (unquoted filename*)
  res.setHeader("Content-Disposition", `attachment; filename="${item.asciiName}"; filename*=UTF-8''${encodeURIComponent(item.filename)}`);
  res.setHeader("Content-Length", item.buffer.length);
  res.send(item.buffer);
});

// Endpoint 3: Direct POST download (supports form submission & direct fetch)
app.post(["/api/download-docx", "/api/export-docx"], async (req, res) => {
  try {
    const {
      headline = "",
      russianText = "",
      uzbekCyrillic = "",
      uzbekLatin = "",
      script = "cyrillic",
      sourceUrl = ""
    } = req.body;

    const doc = createWordDocument({
      headline,
      russianText,
      uzbekCyrillic,
      uzbekLatin,
      script,
      sourceUrl
    });

    const buffer = await Packer.toBuffer(doc);
    const cleanPrefix = (headline || "Maqola")
      .split(/\s+/)
      .slice(0, 4)
      .join(" ")
      .replace(/[^\p{L}\p{N}\s_-]/gu, "")
      .trim() || "Maqola";
    const filename = `${cleanPrefix}.docx`;
    const asciiName = `${toSafeAsciiFilename(cleanPrefix)}.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.setHeader("Content-Length", buffer.length);
    res.send(buffer);
  } catch (err: any) {
    console.error("Error in docx export:", err);
    res.status(500).json({ error: "Word ҳужжатини яратишда хатолик: " + err.message });
  }
});

// Vite middleware & Static hosting
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, HOST, () => {
    console.log(`Дастур ишга тушди: http://localhost:${PORT}`);
  });
}

startServer();
