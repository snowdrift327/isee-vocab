import os
import json
import random
from datetime import datetime
from pathlib import Path

import pandas as pd
from anthropic import Anthropic
from dotenv import load_dotenv

# ============ 配置 ============
load_dotenv()
client = Anthropic()

EXCEL_FILE = Path("isee_lower_level_words.xlsx")
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

INDEX_FILE = Path("index.html")
MISTAKES_FILE = Path("mistakes.html")
HISTORY_FILE = Path("history.html")
WRONG_REVIEW_FILE = Path("wrong-review.html")

NUM_SYNONYM = 14
NUM_SENTENCE = 6
TOTAL_QUESTIONS = NUM_SYNONYM + NUM_SENTENCE

# ============ 读取 Excel 词库 ============
def load_word_pool():
    """从 Excel 读取所有单词，扁平化去重"""
    df = pd.read_excel(EXCEL_FILE)
    word_set = set()
    # 列头本身是主题词，也加入
    for col in df.columns:
        word_set.add(str(col).strip())
        for word in df[col].dropna():
            cleaned = str(word).strip()
            if cleaned:
                word_set.add(cleaned)
    return sorted(word_set)

word_pool = load_word_pool()
print(f"📚 词库加载完成，共 {len(word_pool)} 个独特单词\n")

# ============ 随机选词 ============
selected = random.sample(word_pool, TOTAL_QUESTIONS)
synonym_words = selected[:NUM_SYNONYM]
sentence_words = selected[NUM_SYNONYM:]

print(f"🎲 本次抽取 20 个词：")
print(f"   同义词题（{NUM_SYNONYM} 个）：{', '.join(synonym_words)}")
print(f"   句子完成题（{NUM_SENTENCE} 个）：{', '.join(sentence_words)}\n")

# ============ 构建 Prompt ============
prompt = f"""You are an experienced ISEE Lower Level test question writer with deep knowledge of the actual exam format and difficulty calibration. Your task is to create questions that match the EXACT difficulty of the real ISEE Lower Level vocabulary section.

【CRITICAL DIFFICULTY CALIBRATION - READ CAREFULLY】

ISEE Lower Level is for students applying to Grades 5-7. The vocabulary section is HARDER than school vocabulary tests because:
- Target words are at advanced middle-school to high-school prep level
- Distractors are chosen to test PRECISE word knowledge, not gross errors
- Sentence completion sentences use complex grammatical structures (subordinate clauses, contrasts, abstract concepts)
- Sentence contexts require INFERENCE, not just word matching

【REAL ISEE LOWER LEVEL EXAMPLE - SYNONYM】

PRUDENT most nearly means
(A) wealthy
(B) cautious
(C) friendly
(D) intelligent

Note: All 4 options are common positive-trait adjectives. The student must know the PRECISE meaning of "prudent" — not just guess based on positive connotation. "Intelligent" is a tempting wrong answer because prudent people are often intelligent, but that's not the definition.

【REAL ISEE LOWER LEVEL EXAMPLE - SENTENCE COMPLETION】

Despite the team's initial enthusiasm, their progress became increasingly ______ as the project's complexity overwhelmed their resources.
(A) rapid
(B) sporadic
(C) systematic
(D) deliberate

Note: Uses "Despite" + "increasingly" structure. Context is ABSTRACT (project management, complexity, resources), not concrete (detective, kitchen, etc.). Distractors are all adverbs of work-pace, requiring precise contextual understanding.

【SYNONYM QUESTIONS - generate {NUM_SYNONYM} questions for these words】
{json.dumps(synonym_words)}

REQUIREMENTS for synonym questions:
1. Stem format EXACTLY: "[WORD] most nearly means" (WORD in caps)
2. Correct answer: a precise single-word synonym (not a paraphrase)
3. Three distractors MUST follow these rules:
   - Same part of speech as target (adjective→adjectives only)
   - Same general semantic field (e.g., if target is a positive trait, at least 2 distractors should also be positive traits)
   - At least ONE distractor must be semantically TEMPTING (commonly confused with the target, related but not synonymous)
   - NO obviously wrong choices like "tall" for "meticulous"
4. Difficulty of options: Use vocabulary at the same level as the target word, NOT simple Grade 3 words
5. Vary correct answer position (don't cluster at B or C)

【SENTENCE COMPLETION QUESTIONS - generate {NUM_SENTENCE} questions for these words】
{json.dumps(sentence_words)}

REQUIREMENTS for sentence completion:
1. Sentences MUST use at least ONE of these complex structures:
   - Contrast/concession: "Although...", "Despite...", "Whereas...", "However..."
   - Causation: "Because...", "Since...", "Given that..."
   - Subordinate clause: "..., which...", "When..., the..."
   - Comparison: "more...than...", "as...as..."
2. Sentences MUST involve ABSTRACT concepts when possible (ideas, relationships, qualities, processes) — NOT concrete simple scenarios like "kitchen", "playground", "puppy"
3. Sentence length: 15-25 words (longer than elementary school sentences)
4. The blank should require READING THE WHOLE SENTENCE to fill correctly — not guessable from one nearby word
5. Distractors must be:
   - Grammatically correct in the blank
   - Plausibly fitting the surface meaning, WRONG on deeper inference
   - Same part of speech and similar register as target
6. Use exactly ______ (6 underscores) for the blank

【ANTI-SIMPLIFICATION CHECKLIST】

Before finalizing each question, verify:
□ Synonym: Are distractors at the same difficulty level as the target word? (If target is "meticulous", distractors should NOT be "tall/happy/fast")
□ Sentence: Does the sentence use a complex structure with at least one subordinating word?
□ Sentence: Could a student get the answer right WITHOUT reading the whole sentence? (If yes, rewrite)
□ Sentence: Is the context abstract (idea, relationship, quality) rather than concrete (object, place)?

【OUTPUT FORMAT - STRICT JSON, no markdown, no commentary】

{{
  "questions": [
    {{
      "type": "synonym",
      "word": "METICULOUS",
      "definition": "Showing great attention to detail; very careful and precise.",
      "stem": "METICULOUS most nearly means",
      "options": ["thorough", "rigorous", "scrupulous", "diligent"],
      "correct_index": 2
    }},
    {{
      "type": "sentence",
      "word": "meticulous",
      "definition": "Showing great attention to detail; very careful and precise.",
      "stem": "Although her colleagues worked quickly, Maria's ______ approach to research often uncovered errors that others had overlooked.",
      "options": ["careless", "meticulous", "hesitant", "ambitious"],
      "correct_index": 1
    }}
  ]
}}

FINAL REQUIREMENTS:
- Output exactly {TOTAL_QUESTIONS} questions
- First {NUM_SYNONYM} are synonym, last {NUM_SENTENCE} are sentence completion
- Word field: UPPERCASE for synonyms, lowercase for sentence completion
- correct_index: 0/1/2/3, vary the distribution
- Output ONLY the JSON, no markdown code blocks
"""

# ============ 调用 Claude ============
print("⏳ 正在调用 Claude 生成题目...\n")

message = client.messages.create(
    model="claude-sonnet-4-5",
    max_tokens=8192,
    messages=[{"role": "user", "content": prompt}]
)

response_text = message.content[0].text.strip()
if response_text.startswith("```"):
    response_text = response_text.split("```")[1]
    if response_text.startswith("json"):
        response_text = response_text[4:]
    response_text = response_text.strip()

data = json.loads(response_text)
questions = data["questions"]

if len(questions) != TOTAL_QUESTIONS:
    print(f"⚠️  警告：AI 返回了 {len(questions)} 道题，期望 {TOTAL_QUESTIONS} 道")

print(f"✅ 成功生成 {len(questions)} 道题\n")

# ============ 保存本次会话数据 ============
timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
session_file = DATA_DIR / f"{timestamp}.json"
with open(session_file, "w", encoding="utf-8") as f:
    json.dump({
        "timestamp": timestamp,
        "words": selected,
        "questions": questions
    }, f, ensure_ascii=False, indent=2)
print(f"💾 会话数据已保存：{session_file}\n")

# ============ 生成 index.html ============
def render_index_html(questions, timestamp):
    """生成 ISEE 风格的做题页面"""

    # 把题目数据嵌入到 JS 中（用 json.dumps 保证安全）
    questions_json = json.dumps(questions, ensure_ascii=False)

    # 生成题目 HTML
    questions_html = ""
    for i, q in enumerate(questions):
        # 题号 + 题目类型标记
        q_type_label = "Synonym" if q["type"] == "synonym" else "Sentence Completion"

        # 题干：同义词题需要把单词显示成 SMALL CAPS 样式
        if q["type"] == "synonym":
            stem_display = q["stem"].replace(q["word"], f'<strong class="target-word">{q["word"]}</strong>')
        else:
            stem_display = q["stem"].replace("______", '<span class="blank">______</span>')

        # 4 个选项
        options_html = ""
        letters = ['A', 'B', 'C', 'D']
        for j, opt in enumerate(q["options"]):
            options_html += f'''
            <label class="option" data-q="{i}" data-opt="{j}">
                <input type="radio" name="q{i}" value="{j}">
                <span class="letter">{letters[j]}</span>
                <span class="opt-text">{opt}</span>
            </label>'''

        questions_html += f'''
        <div class="question" data-q-index="{i}">
            <div class="q-header">
                <span class="q-number">{i+1}.</span>
                <span class="q-type">{q_type_label}</span>
            </div>
            <div class="q-stem">{stem_display}</div>
            <div class="q-options">{options_html}</div>
        </div>'''

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ISEE Lower Level — Vocabulary Practice</title>
<style>
/* Start 遮罩 */
.start-overlay {{
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(255, 255, 255, 0.98);
    z-index: 9999;
    display: flex;
    align-items: center;
    justify-content: center;
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
}}
.start-card {{
    background: white;
    max-width: 480px;
    width: 90%;
    padding: 40px 32px;
    border: 3px double #1a4d8f;
    border-radius: 8px;
    text-align: center;
    box-shadow: 0 8px 32px rgba(0,0,0,0.15);
    font-family: "Times New Roman", Georgia, serif;
}}
.start-card h2 {{
    font-size: 26px;
    color: #1a4d8f;
    margin: 0 0 8px;
    letter-spacing: 1px;
}}
.start-card .start-subtitle {{
    color: #666;
    font-size: 14px;
    font-style: italic;
    margin-bottom: 24px;
}}
.start-card .instructions {{
    text-align: left;
    background: #f5f5f0;
    padding: 16px 20px;
    border-radius: 6px;
    margin: 20px 0 28px;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
    line-height: 1.7;
    color: #333;
}}
.start-card .instructions strong {{
    color: #1a4d8f;
    display: block;
    margin-bottom: 6px;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 1px;
}}
.start-card .instructions ul {{
    margin: 0;
    padding-left: 20px;
}}
.start-card .instructions li {{
    margin: 4px 0;
}}
.start-btn {{
    background: #1a4d8f;
    color: white;
    border: none;
    padding: 16px 56px;
    font-size: 17px;
    font-weight: bold;
    border-radius: 4px;
    cursor: pointer;
    font-family: "Helvetica Neue", Arial, sans-serif;
    letter-spacing: 2px;
    transition: background 0.2s;
}}
.start-btn:hover {{
    background: #143a6b;
}}
.start-card .meta-info {{
    margin-top: 20px;
    font-size: 12px;
    color: #888;
    font-family: Arial, sans-serif;
}}
.hidden-until-start {{
    visibility: hidden;
}}
.start-overlay.hidden {{
    display: none;
}}
* {{ box-sizing: border-box; }}
body {{
    font-family: "Times New Roman", Georgia, serif;
    max-width: 760px;
    margin: 0 auto;
    padding: 24px 20px 80px;
    background: #ffffff;
    color: #1a1a1a;
    line-height: 1.6;
    font-size: 16px;
}}
.test-header {{
    border-bottom: 3px double #333;
    padding-bottom: 16px;
    margin-bottom: 28px;
}}
h1 {{
    font-size: 24px;
    margin: 0 0 6px;
    color: #1a1a1a;
    font-weight: bold;
    letter-spacing: 0.5px;
}}
.subtitle {{ font-size: 13px; color: #555; font-style: italic; }}
.meta-bar {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 14px;
    padding: 10px 16px;
    background: #f5f5f0;
    border: 1px solid #ddd;
    border-radius: 4px;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
}}
.timer {{
    font-family: "Courier New", monospace;
    font-size: 17px;
    font-weight: bold;
    color: #1a4d8f;
}}
.progress {{ color: #555; }}

.question {{
    margin-bottom: 28px;
    padding-bottom: 24px;
    border-bottom: 1px solid #eee;
}}
.question:last-child {{ border-bottom: none; }}
.q-header {{ margin-bottom: 8px; }}
.q-number {{
    font-size: 18px;
    font-weight: bold;
    color: #1a1a1a;
    margin-right: 10px;
}}
.q-type {{
    font-size: 12px;
    color: #888;
    font-style: italic;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
.q-stem {{
    font-size: 16px;
    margin-bottom: 14px;
    line-height: 1.6;
}}
.target-word {{
    font-variant: small-caps;
    letter-spacing: 1px;
    font-weight: bold;
}}
.blank {{
    color: #1a4d8f;
    font-weight: bold;
    letter-spacing: 2px;
}}
.q-options {{
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-left: 20px;
}}
.option {{
    display: flex;
    align-items: flex-start;
    padding: 8px 14px;
    border: 1px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    transition: background 0.15s;
    font-size: 15px;
}}
.option:hover {{ background: #f5f5f0; }}
.option input[type="radio"] {{
    margin-right: 10px;
    margin-top: 4px;
    cursor: pointer;
}}
.letter {{
    font-weight: bold;
    color: #555;
    margin-right: 8px;
    min-width: 18px;
}}
.opt-text {{ flex: 1; }}

/* 提交后的状态样式 */
.option.user-correct {{
    background: #e8f5e9;
    border-color: #4caf50;
}}
.option.user-wrong {{
    background: #ffebee;
    border-color: #f44336;
}}
.option.show-correct {{
    background: #e8f5e9;
    border-color: #4caf50;
    border-style: dashed;
}}

.submit-section {{
    margin-top: 40px;
    text-align: center;
    padding-top: 24px;
    border-top: 3px double #333;
}}
#submit-btn {{
    background: #1a4d8f;
    color: white;
    border: none;
    padding: 12px 48px;
    font-size: 16px;
    font-weight: bold;
    border-radius: 4px;
    cursor: pointer;
    font-family: "Helvetica Neue", Arial, sans-serif;
    letter-spacing: 1px;
    transition: background 0.2s;
}}
#submit-btn:hover {{ background: #143a6b; }}
#submit-btn:disabled {{
    background: #aaa;
    cursor: not-allowed;
}}

.new-test-section {{
    text-align: center;
    margin-bottom: 20px;
    padding-bottom: 18px;
    border-bottom: 1px solid #ddd;
}}
.start-new-btn {{
    background: #2e7d32;
    color: white;
    border: none;
    padding: 12px 40px;
    font-size: 15px;
    font-weight: bold;
    border-radius: 6px;
    cursor: pointer;
    font-family: "Helvetica Neue", Arial, sans-serif;
    letter-spacing: 1px;
    transition: background 0.2s;
}}
.start-new-btn:hover {{
    background: #1b5e20;
}}

/* 结果区 */
#results {{
    display: none;
    margin-top: 30px;
    padding: 24px;
    background: #f9f9f7;
    border: 2px solid #1a4d8f;
    border-radius: 6px;
    font-family: "Helvetica Neue", Arial, sans-serif;
}}
#results.show {{ display: block; }}
.score-summary {{
    display: flex;
    justify-content: space-around;
    text-align: center;
    margin-bottom: 20px;
    flex-wrap: wrap;
    gap: 14px;
}}
.score-item {{ flex: 1; min-width: 120px; }}
.score-value {{
    font-size: 32px;
    font-weight: bold;
    color: #1a4d8f;
    display: block;
}}
.score-label {{
    font-size: 12px;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 4px;
}}

.mistakes-list {{ margin-top: 20px; }}
.mistakes-list h3 {{
    font-size: 16px;
    color: #c62828;
    margin: 0 0 12px;
    border-bottom: 1px solid #ddd;
    padding-bottom: 8px;
}}
.mistake-item {{
    background: white;
    padding: 14px 16px;
    margin-bottom: 10px;
    border-left: 4px solid #f44336;
    border-radius: 0 4px 4px 0;
}}
.mistake-word {{
    font-weight: bold;
    font-size: 16px;
    color: #1a1a1a;
    margin-bottom: 4px;
}}
.mistake-def {{
    color: #444;
    font-size: 14px;
    font-style: italic;
    margin-bottom: 6px;
}}
.mistake-detail {{
    font-size: 13px;
    color: #666;
    line-height: 1.5;
}}
.action-links {{
    margin: 20px 0;
    padding: 16px 0;
    border-top: 1px solid #ddd;
    border-bottom: 1px solid #ddd;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
}}
.action-links a {{
    display: block;
    text-align: center;
    padding: 12px 16px;
    background: #1a4d8f;
    color: white;
    text-decoration: none;
    border-radius: 4px;
    font-size: 14px;
    font-weight: 500;
}}
.action-links a:hover {{ background: #143a6b; }}
@media (max-width: 480px) {{
    .action-links {{ grid-template-columns: 1fr; }}
}}

footer {{
    text-align: center;
    margin-top: 50px;
    color: #999;
    font-size: 12px;
    font-family: Arial, sans-serif;
}}

@media (max-width: 600px) {{
    body {{ padding: 16px 14px 60px; font-size: 15px; }}
    h1 {{ font-size: 20px; }}
    .meta-bar {{ flex-direction: column; gap: 6px; align-items: flex-start; }}
    .q-options {{ margin-left: 8px; }}
    .option {{ padding: 8px 10px; font-size: 14px; }}
    .score-value {{ font-size: 26px; }}
}}
</style>
</head>
<body>
<div id="start-overlay" class="start-overlay">
    <div class="start-card">
        <h2>ISEE LOWER LEVEL</h2>
        <div class="start-subtitle">Vocabulary Practice Test</div>

        <div class="instructions">
            <strong>Instructions</strong>
            <ul>
                <li><strong>{len(questions)} questions</strong> ({NUM_SYNONYM} synonym + {NUM_SENTENCE} sentence completion)</li>
                <li>Timer starts when you click <strong>START</strong></li>
                <li>Select one answer for each question</li>
                <li>Click <strong>SUBMIT TEST</strong> when finished</li>
                <li>Wrong answers will be saved to your Mistakes Book</li>
            </ul>
        </div>

        <button class="start-btn" onclick="startTest()">START</button>

        <div class="meta-info">Generated: {timestamp}</div>
    </div>
</div>
<div class="test-header">
    <h1>ISEE LOWER LEVEL</h1>
    <div class="subtitle">Vocabulary Practice — Synonyms &amp; Sentence Completion</div>
    <div class="meta-bar">
        <span class="timer">Time: <span id="timer-display">00:00</span></span>
        <span class="progress">Questions: <span id="progress-count">0</span> / {len(questions)}</span>
        <span style="color:#888;font-size:12px;">{timestamp}</span>
    </div>
</div>

<form id="quiz-form" onsubmit="return false;">
{questions_html}
</form>

<div class="submit-section">
    <button id="submit-btn" onclick="submitQuiz()">SUBMIT TEST</button>
</div>

<div id="results"></div>

<footer>
    <a href="mistakes.html" style="color:#888;">📖 View Mistakes Book</a>
    &nbsp;|&nbsp;
    <a href="history.html" style="color:#888;">📊 View History</a>
    &nbsp;|&nbsp;
    Generated by Claude
</footer>

<script>
// 检查是否处于"重做错题"模式
// === 变量声明（必须放在最前面） ===
let startTime = null;
let submitted = false;
let testStarted = false;
let QUESTIONS_DATA;

// === 优先级 1：检查是否有上次提交的状态需要恢复 ===
const savedFormHTML = sessionStorage.getItem('submitted_form_html');

// === 检查是否处于"重做错题"模式 ===
const isRetakeMode = sessionStorage.getItem('retake_mode') === '1';

if (savedFormHTML) {{
    // 恢复上次提交后的完整结果页
    QUESTIONS_DATA = {questions_json};
    submitted = true;

    document.getElementById('start-overlay').classList.add('hidden');
    document.getElementById('quiz-form').innerHTML = savedFormHTML;
    document.getElementById('results').innerHTML = sessionStorage.getItem('submitted_results_html');
    document.getElementById('results').classList.add('show');
    document.getElementById('timer-display').textContent = sessionStorage.getItem('submitted_timer');
    document.getElementById('progress-count').textContent = sessionStorage.getItem('submitted_progress');
    document.getElementById('submit-btn').disabled = true;
    document.getElementById('submit-btn').textContent = 'SUBMITTED';
}} else if (isRetakeMode) {{
    QUESTIONS_DATA = JSON.parse(sessionStorage.getItem('retake_questions') || '[]');
    sessionStorage.removeItem('retake_mode');
    sessionStorage.removeItem('retake_questions');

    if (QUESTIONS_DATA.length === 0) {{
        document.body.innerHTML = '<div style="padding:60px;text-align:center;font-family:Arial;"><h2>No wrong questions to retake.</h2><a href="index.html">← Back to start</a></div>';
    }} else {{
        // 重做模式：自动开始，跳过 Start 遮罩
        document.getElementById('start-overlay').classList.add('hidden');
        testStarted = true;
        startTime = Date.now();
        // 重新渲染题目
        rerenderQuestions(QUESTIONS_DATA);
    }}
}} else {{
    QUESTIONS_DATA = {questions_json};
}}

function startTest() {{
    testStarted = true;
    startTime = Date.now();
    document.getElementById('start-overlay').classList.add('hidden');
}}

function rerenderQuestions(questions) {{
    // 更新计数
    document.querySelector('.progress').innerHTML = 'Questions: <span id="progress-count">0</span> / ' + questions.length;

    // 添加重做模式标记
    const headerSubtitle = document.querySelector('.subtitle');
    headerSubtitle.innerHTML = '🔁 RETAKING WRONG QUESTIONS — ' + questions.length + ' question(s)';
    headerSubtitle.style.color = '#c62828';
    headerSubtitle.style.fontWeight = 'bold';

    // 清空原题目，重新渲染
    const form = document.getElementById('quiz-form');
    form.innerHTML = '';

    questions.forEach((q, i) => {{
        const qDiv = document.createElement('div');
        qDiv.className = 'question';
        qDiv.dataset.qIndex = i;

        let stemDisplay;
        if (q.type === 'synonym') {{
            stemDisplay = q.stem.replace(q.word, '<strong class="target-word">' + q.word + '</strong>');
        }} else {{
            stemDisplay = q.stem.replace(/______/g, '<span class="blank">______</span>');
        }}

        const typeLabel = q.type === 'synonym' ? 'Synonym' : 'Sentence Completion';

        let optionsHtml = '';
        const letters = ['A', 'B', 'C', 'D'];
        q.options.forEach((opt, j) => {{
            optionsHtml += '<label class="option" data-q="' + i + '" data-opt="' + j + '">' +
                '<input type="radio" name="q' + i + '" value="' + j + '">' +
                '<span class="letter">' + letters[j] + '</span>' +
                '<span class="opt-text">' + opt + '</span></label>';
        }});

        qDiv.innerHTML = '<div class="q-header"><span class="q-number">' + (i+1) + '.</span>' +
            '<span class="q-type">' + typeLabel + '</span></div>' +
            '<div class="q-stem">' + stemDisplay + '</div>' +
            '<div class="q-options">' + optionsHtml + '</div>';

        form.appendChild(qDiv);
    }});

    // 重新绑定 change 事件
    document.querySelectorAll('input[type="radio"]').forEach(input => {{
        input.addEventListener('change', updateProgress);
    }});
}}

function updateProgress() {{
    const answered = new Set();
    document.querySelectorAll('input[type="radio"]:checked').forEach(r => {{
        answered.add(r.name);
    }});
    document.getElementById('progress-count').textContent = answered.size;
}}

// 计时器
function updateTimer() {{
    if (!testStarted || submitted) return;
    const elapsed = Math.floor((Date.now() - startTime) / 1000);
    const mm = String(Math.floor(elapsed / 60)).padStart(2, '0');
    const ss = String(elapsed % 60).padStart(2, '0');
    document.getElementById('timer-display').textContent = mm + ':' + ss;
}}
setInterval(updateTimer, 1000);

// 进度计数（原题模式）
document.querySelectorAll('input[type="radio"]').forEach(input => {{
    input.addEventListener('change', updateProgress);
}});

function submitQuiz() {{
    if (submitted) return;

    const userAnswers = {{}};
    let answeredCount = 0;
    QUESTIONS_DATA.forEach((q, i) => {{
        const selected = document.querySelector('input[name="q' + i + '"]:checked');
        if (selected) {{
            userAnswers[i] = parseInt(selected.value);
            answeredCount++;
        }} else {{
            userAnswers[i] = -1;
        }}
    }});

    if (answeredCount < QUESTIONS_DATA.length) {{
        if (!confirm('You have ' + (QUESTIONS_DATA.length - answeredCount) + ' unanswered question(s). Submit anyway?')) {{
            return;
        }}
    }}

    submitted = true;
    const totalTime = startTime ? Math.floor((Date.now() - startTime) / 1000) : 0;

    let correctCount = 0;
    const mistakes = [];
    const wrongQuestions = [];  // 保留完整题目数据，用于重做

    QUESTIONS_DATA.forEach((q, i) => {{
        const userAns = userAnswers[i];
        const correctAns = q.correct_index;

        const allOpts = document.querySelectorAll('label.option[data-q="' + i + '"]');
        allOpts.forEach(opt => {{
            const optIdx = parseInt(opt.dataset.opt);
            opt.style.pointerEvents = 'none';
            const radio = opt.querySelector('input');
            radio.disabled = true;

            if (optIdx === correctAns) {{
                if (userAns === correctAns) {{
                    opt.classList.add('user-correct');
                }} else {{
                    opt.classList.add('show-correct');
                }}
            }} else if (optIdx === userAns) {{
                opt.classList.add('user-wrong');
            }}
        }});

        if (userAns === correctAns) {{
            correctCount++;
        }} else {{
            mistakes.push({{
                word: q.word,
                definition: q.definition,
                type: q.type,
                stem: q.stem,
                options: q.options,
                correct_index: q.correct_index,
                user_index: userAns,
                date: new Date().toISOString().split('T')[0]
            }});
            wrongQuestions.push(q);
        }}
    }});

    // 永久错题本（localStorage）
    const existing = JSON.parse(localStorage.getItem('isee_mistakes') || '[]');
    mistakes.forEach(m => {{
        const isDup = existing.some(e => e.word === m.word && e.stem === m.stem);
        if (!isDup) existing.push(m);
    }});
    localStorage.setItem('isee_mistakes', JSON.stringify(existing));

    // 本次错题（sessionStorage，用于重做）
    sessionStorage.setItem('last_wrong_questions', JSON.stringify(wrongQuestions));

    // 渲染结果
    // 保存本次会话到历史记录
    const sessionRecord = {{
        date: new Date().toISOString().split('T')[0],
        timestamp: new Date().toISOString(),
        type: 'new',
        total: QUESTIONS_DATA.length,
        correct: correctCount,
        accuracy: Math.round((correctCount / QUESTIONS_DATA.length) * 100),
        duration_sec: totalTime
    }};
    const sessionsList = JSON.parse(localStorage.getItem('isee_sessions') || '[]');
    sessionsList.push(sessionRecord);
    localStorage.setItem('isee_sessions', JSON.stringify(sessionsList));
    const accuracy = Math.round((correctCount / QUESTIONS_DATA.length) * 100);
    const mm = String(Math.floor(totalTime / 60)).padStart(2, '0');
    const ss = String(totalTime % 60).padStart(2, '0');

    let mistakesHtml = '';
    if (mistakes.length > 0) {{
        mistakesHtml = '<div class="mistakes-list"><h3>📝 Review These Words (' + mistakes.length + ')</h3>';
        mistakes.forEach(m => {{
            const correctText = m.options[m.correct_index];
            const userText = m.user_index >= 0 ? m.options[m.user_index] : '(unanswered)';
            mistakesHtml += '<div class="mistake-item">' +
                '<div class="mistake-word">' + m.word + '</div>' +
                '<div class="mistake-def">' + m.definition + '</div>' +
                '<div class="mistake-detail">' +
                '<strong>Correct:</strong> ' + correctText + '<br>' +
                '<strong>Your answer:</strong> ' + userText +
                '</div></div>';
        }});
        mistakesHtml += '</div>';
    }} else {{
        mistakesHtml = '<div style="text-align:center;padding:20px;color:#2e7d32;font-size:18px;font-weight:bold;">🎉 Perfect Score! All correct!</div>';
    }}

    // 重做按钮（只有有错题时才显示）
    const retakeBtn = wrongQuestions.length > 0
        ? '<a href="javascript:retakeWrong()">🔁 Retake Wrong Questions (' + wrongQuestions.length + ')</a>'
        : '';

    document.getElementById('results').innerHTML =
        '<div class="new-test-section">' +
        '<button class="start-new-btn" onclick="startNewTest()">🔄 Start New Test</button>' +
        '</div>' +
        '<div class="score-summary">' +
        '<div class="score-item"><span class="score-value">' + correctCount + '/' + QUESTIONS_DATA.length + '</span><span class="score-label">Score</span></div>' +
        '<div class="score-item"><span class="score-value">' + accuracy + '%</span><span class="score-label">Accuracy</span></div>' +
        '<div class="score-item"><span class="score-value">' + mm + ':' + ss + '</span><span class="score-label">Time</span></div>' +
        '</div>' +
       
        '<div class="action-links">' +
        retakeBtn +
        '<a href="mistakes.html">📖 View All Mistakes</a>' +
        '<a href="wrong-review.html">🔁 Review Mistakes</a>' +
                '<a href="history.html">📊 View Progress</a>' +
        '</div>' +
        mistakesHtml;
        

    document.getElementById('results').classList.add('show');
    document.getElementById('submit-btn').disabled = true;
    document.getElementById('submit-btn').textContent = 'SUBMITTED';

    document.getElementById('results').scrollIntoView({{behavior: 'smooth'}});

    // 保存完整状态到 sessionStorage（用于跨页面返回时恢复）
    sessionStorage.setItem('submitted_form_html', document.getElementById('quiz-form').innerHTML);
    sessionStorage.setItem('submitted_results_html', document.getElementById('results').innerHTML);
    sessionStorage.setItem('submitted_timer', document.getElementById('timer-display').textContent);
    sessionStorage.setItem('submitted_progress', document.getElementById('progress-count').textContent);
}}

function clearSubmittedState() {{
    sessionStorage.removeItem('submitted_form_html');
    sessionStorage.removeItem('submitted_results_html');
    sessionStorage.removeItem('submitted_timer');
    sessionStorage.removeItem('submitted_progress');
}}

function startNewTest() {{
    clearSubmittedState();
    location.reload();
}}

function retakeWrong() {{
    const wrongQuestions = sessionStorage.getItem('last_wrong_questions');
    if (!wrongQuestions || JSON.parse(wrongQuestions).length === 0) {{
        alert('No wrong questions to retake.');
        return;
    }}
    clearSubmittedState();
    sessionStorage.setItem('retake_mode', '1');
    sessionStorage.setItem('retake_questions', wrongQuestions);
    location.reload();
}}

</script>

</body>
</html>'''

# ============ 生成 wrong-review.html ============
def render_wrong_review_html():
    """错题集中复习页面（从 localStorage 读错题）"""
    return '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ISEE Vocabulary — Mistakes Review</title>
<style>
* { box-sizing: border-box; }
body {
    font-family: "Times New Roman", Georgia, serif;
    max-width: 760px;
    margin: 0 auto;
    padding: 24px 20px 80px;
    background: #ffffff;
    color: #1a1a1a;
    line-height: 1.6;
    font-size: 16px;
}
.start-overlay {
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(255, 255, 255, 0.98);
    z-index: 9999;
    display: flex;
    align-items: center;
    justify-content: center;
    backdrop-filter: blur(8px);
}
.start-overlay.hidden { display: none; }
.start-card {
    background: white;
    max-width: 480px;
    width: 90%;
    padding: 40px 32px;
    border: 3px double #c62828;
    border-radius: 8px;
    text-align: center;
    box-shadow: 0 8px 32px rgba(0,0,0,0.15);
}
.start-card h2 {
    font-size: 26px;
    color: #c62828;
    margin: 0 0 8px;
    letter-spacing: 1px;
}
.start-card .start-subtitle {
    color: #666;
    font-size: 14px;
    font-style: italic;
    margin-bottom: 24px;
}
.instructions {
    text-align: left;
    background: #fdf6f6;
    padding: 16px 20px;
    border-radius: 6px;
    margin: 20px 0 28px;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
    line-height: 1.7;
    color: #333;
}
.instructions strong {
    color: #c62828;
    display: block;
    margin-bottom: 6px;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 1px;
}
.instructions ul { margin: 0; padding-left: 20px; }
.instructions li { margin: 4px 0; }
.start-btn {
    background: #c62828;
    color: white;
    border: none;
    padding: 16px 56px;
    font-size: 17px;
    font-weight: bold;
    border-radius: 4px;
    cursor: pointer;
    font-family: "Helvetica Neue", Arial, sans-serif;
    letter-spacing: 2px;
    transition: background 0.2s;
}
.start-btn:hover { background: #9e1f1f; }
.start-btn:disabled { background: #aaa; cursor: not-allowed; }
.empty-info {
    padding: 30px 20px;
    text-align: center;
}
.meta-info {
    margin-top: 20px;
    font-size: 12px;
    color: #888;
    font-family: Arial, sans-serif;
}

.test-header {
    border-bottom: 3px double #c62828;
    padding-bottom: 16px;
    margin-bottom: 28px;
}
h1 { font-size: 24px; margin: 0 0 6px; color: #c62828; font-weight: bold; }
.page-subtitle { font-size: 13px; color: #555; font-style: italic; }
.meta-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 14px;
    padding: 10px 16px;
    background: #fdf6f6;
    border: 1px solid #f0d6d6;
    border-radius: 4px;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
}
.timer {
    font-family: "Courier New", monospace;
    font-size: 17px;
    font-weight: bold;
    color: #c62828;
}

.question {
    margin-bottom: 28px;
    padding-bottom: 24px;
    border-bottom: 1px solid #eee;
}
.question:last-child { border-bottom: none; }
.q-header { margin-bottom: 8px; }
.q-number { font-size: 18px; font-weight: bold; margin-right: 10px; }
.q-type {
    font-size: 12px;
    color: #888;
    font-style: italic;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.q-stem { font-size: 16px; margin-bottom: 14px; line-height: 1.6; }
.target-word {
    font-variant: small-caps;
    letter-spacing: 1px;
    font-weight: bold;
}
.blank { color: #c62828; font-weight: bold; letter-spacing: 2px; }
.q-options {
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-left: 20px;
}
.option {
    display: flex;
    align-items: flex-start;
    padding: 8px 14px;
    border: 1px solid transparent;
    border-radius: 4px;
    cursor: pointer;
    font-size: 15px;
}
.option:hover { background: #fdf6f6; }
.option input[type="radio"] { margin-right: 10px; margin-top: 4px; }
.letter { font-weight: bold; color: #555; margin-right: 8px; min-width: 18px; }
.option.user-correct { background: #e8f5e9; border-color: #4caf50; }
.option.user-wrong { background: #ffebee; border-color: #f44336; }
.option.show-correct { background: #e8f5e9; border-color: #4caf50; border-style: dashed; }

.submit-section {
    margin-top: 40px;
    text-align: center;
    padding-top: 24px;
    border-top: 3px double #c62828;
}
#submit-btn {
    background: #c62828;
    color: white;
    border: none;
    padding: 12px 48px;
    font-size: 16px;
    font-weight: bold;
    border-radius: 4px;
    cursor: pointer;
    font-family: "Helvetica Neue", Arial, sans-serif;
    letter-spacing: 1px;
}
#submit-btn:hover { background: #9e1f1f; }
#submit-btn:disabled { background: #aaa; cursor: not-allowed; }

#results {
    display: none;
    margin-top: 30px;
    padding: 24px;
    background: #fdf6f6;
    border: 2px solid #c62828;
    border-radius: 6px;
    font-family: "Helvetica Neue", Arial, sans-serif;
}
#results.show { display: block; }
.score-summary {
    display: flex;
    justify-content: space-around;
    text-align: center;
    margin-bottom: 20px;
    flex-wrap: wrap;
    gap: 14px;
}
.score-item { flex: 1; min-width: 120px; }
.score-value { font-size: 32px; font-weight: bold; color: #c62828; display: block; }
.score-label {
    font-size: 12px;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 4px;
}

.mastered-list, .still-wrong-list { margin-top: 20px; }
.mastered-list h3, .still-wrong-list h3 {
    font-size: 16px;
    margin: 0 0 12px;
    border-bottom: 1px solid #ddd;
    padding-bottom: 8px;
}
.mastered-list h3 { color: #2e7d32; }
.still-wrong-list h3 { color: #c62828; }
.mastered-item {
    background: #e8f5e9;
    padding: 10px 14px;
    margin-bottom: 6px;
    border-left: 4px solid #4caf50;
    border-radius: 0 4px 4px 0;
    font-size: 14px;
}
.mistake-item {
    background: white;
    padding: 12px 16px;
    margin-bottom: 8px;
    border-left: 4px solid #f44336;
    border-radius: 0 4px 4px 0;
}
.mistake-word { font-weight: bold; font-size: 15px; margin-bottom: 4px; }
.mistake-def { color: #666; font-size: 13px; font-style: italic; }

.action-links {
    margin: 20px 0;
    padding: 16px 0;
    border-top: 1px solid #ddd;
    border-bottom: 1px solid #ddd;
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 10px;
}
.action-links a {
    display: block;
    text-align: center;
    padding: 12px 16px;
    background: #c62828;
    color: white;
    text-decoration: none;
    border-radius: 4px;
    font-size: 14px;
    font-weight: 500;
}
.action-links a:hover { background: #9e1f1f; }

@media (max-width: 480px) {
    .action-links { grid-template-columns: 1fr; }
}

@media (max-width: 600px) {
    body { padding: 16px 14px 60px; font-size: 15px; }
    h1 { font-size: 20px; }
    .meta-bar { flex-direction: column; gap: 6px; align-items: flex-start; }
    .q-options { margin-left: 8px; }
    .option { padding: 8px 10px; font-size: 14px; }
    .score-value { font-size: 26px; }
}
</style>
</head>
<body>

<div id="start-overlay" class="start-overlay">
    <div class="start-card">
        <h2>🔁 MISTAKES REVIEW</h2>
        <div class="start-subtitle">Practice your previously wrong words</div>
        <div id="review-info"></div>
        <button id="start-btn" class="start-btn" onclick="startReview()" disabled>START REVIEW</button>
        <div class="meta-info">
            <a href="index.html" style="color:#888;">← Back to home</a>
        </div>
    </div>
</div>

<div class="test-header" id="test-header" style="display:none;">
    <h1>🔁 MISTAKES REVIEW</h1>
    <div class="page-subtitle">Review words you've previously gotten wrong</div>
    <div class="meta-bar">
        <span class="timer">Time: <span id="timer-display">00:00</span></span>
        <span>Questions: <span id="progress-count">0</span> / <span id="total-count">0</span></span>
    </div>
</div>

<form id="quiz-form" onsubmit="return false;"></form>

<div class="submit-section" id="submit-section" style="display:none;">
    <button id="submit-btn" onclick="submitReview()">SUBMIT REVIEW</button>
</div>

<div id="results"></div>

<script>
let allMistakes = JSON.parse(localStorage.getItem('isee_mistakes') || '[]');
let reviewQuestions = [];
let startTime = null;
let submitted = false;
let testStarted = false;
let timerInterval = null;

function init() {
    const info = document.getElementById('review-info');
    const btn = document.getElementById('start-btn');

    if (allMistakes.length === 0) {
        info.innerHTML = '<div class="empty-info">' +
            '<p style="font-size:20px;margin:24px 0;color:#2e7d32;">🎉 No mistakes to review!</p>' +
            '<p style="color:#666;">Take a practice test first and your mistakes will appear here.</p>' +
            '</div>';
        btn.style.display = 'none';
    } else {
        const reviewCount = Math.min(10, allMistakes.length);
        info.innerHTML = '<div class="instructions">' +
            '<strong>Review Session</strong>' +
            '<ul>' +
            '<li><strong>' + allMistakes.length + ' word(s)</strong> currently in your Mistakes Book</li>' +
            '<li>This session: practice <strong>' + reviewCount + ' random question(s)</strong></li>' +
            '<li>Words you get RIGHT will be REMOVED from Mistakes Book ✓</li>' +
            '<li>Words you get WRONG will stay for future review</li>' +
            '</ul>' +
            '</div>';
        btn.disabled = false;
    }
}

function startReview() {
    const shuffled = allMistakes.slice().sort(() => Math.random() - 0.5);
    reviewQuestions = shuffled.slice(0, Math.min(10, shuffled.length));

    renderQuestions();
    document.getElementById('start-overlay').classList.add('hidden');
    document.getElementById('test-header').style.display = 'block';
    document.getElementById('submit-section').style.display = 'block';

    testStarted = true;
    startTime = Date.now();
    timerInterval = setInterval(updateTimer, 1000);
}

function renderQuestions() {
    const form = document.getElementById('quiz-form');
    let html = '';
    reviewQuestions.forEach((q, i) => {
        const typeLabel = q.type === 'synonym' ? 'Synonym' : 'Sentence Completion';
        let stemDisplay;
        if (q.type === 'synonym') {
            stemDisplay = q.stem.replace(q.word, '<strong class="target-word">' + q.word + '</strong>');
        } else {
            stemDisplay = q.stem.replace(/______/g, '<span class="blank">______</span>');
        }
        let optionsHtml = '';
        const letters = ['A', 'B', 'C', 'D'];
        q.options.forEach((opt, j) => {
            optionsHtml += '<label class="option" data-q="' + i + '" data-opt="' + j + '">' +
                '<input type="radio" name="q' + i + '" value="' + j + '">' +
                '<span class="letter">' + letters[j] + '</span>' +
                '<span class="opt-text">' + opt + '</span></label>';
        });
        html += '<div class="question" data-q-index="' + i + '">' +
            '<div class="q-header"><span class="q-number">' + (i+1) + '.</span>' +
            '<span class="q-type">' + typeLabel + '</span></div>' +
            '<div class="q-stem">' + stemDisplay + '</div>' +
            '<div class="q-options">' + optionsHtml + '</div>' +
            '</div>';
    });
    form.innerHTML = html;

    document.querySelectorAll('input[type="radio"]').forEach(input => {
        input.addEventListener('change', updateProgress);
    });
    document.getElementById('total-count').textContent = reviewQuestions.length;
}

function updateProgress() {
    const answered = new Set();
    document.querySelectorAll('input[type="radio"]:checked').forEach(r => answered.add(r.name));
    document.getElementById('progress-count').textContent = answered.size;
}

function updateTimer() {
    if (!testStarted || submitted) return;
    const elapsed = Math.floor((Date.now() - startTime) / 1000);
    const mm = String(Math.floor(elapsed / 60)).padStart(2, '0');
    const ss = String(elapsed % 60).padStart(2, '0');
    document.getElementById('timer-display').textContent = mm + ':' + ss;
}

function submitReview() {
    if (submitted) return;

    const userAnswers = {};
    let answeredCount = 0;
    reviewQuestions.forEach((q, i) => {
        const selected = document.querySelector('input[name="q' + i + '"]:checked');
        if (selected) {
            userAnswers[i] = parseInt(selected.value);
            answeredCount++;
        } else {
            userAnswers[i] = -1;
        }
    });

    if (answeredCount < reviewQuestions.length) {
        if (!confirm('You have ' + (reviewQuestions.length - answeredCount) + ' unanswered question(s). Submit anyway?')) {
            return;
        }
    }

    submitted = true;
    clearInterval(timerInterval);
    const totalTime = Math.floor((Date.now() - startTime) / 1000);

    let correctCount = 0;
    const wordsToRemove = [];
    const stillWrong = [];

    reviewQuestions.forEach((q, i) => {
        const userAns = userAnswers[i];
        const correctAns = q.correct_index;

        const allOpts = document.querySelectorAll('label.option[data-q="' + i + '"]');
        allOpts.forEach(opt => {
            const optIdx = parseInt(opt.dataset.opt);
            opt.style.pointerEvents = 'none';
            const radio = opt.querySelector('input');
            radio.disabled = true;
            if (optIdx === correctAns) {
                if (userAns === correctAns) opt.classList.add('user-correct');
                else opt.classList.add('show-correct');
            } else if (optIdx === userAns) {
                opt.classList.add('user-wrong');
            }
        });

        if (userAns === correctAns) {
            correctCount++;
            wordsToRemove.push(q);
        } else {
            stillWrong.push(q);
        }
    });

    // 从错题本移除答对的词
    const newMistakes = allMistakes.filter(m =>
        !wordsToRemove.some(r => r.word === m.word && r.stem === m.stem)
    );
    localStorage.setItem('isee_mistakes', JSON.stringify(newMistakes));

    // 保存到历史成绩
    const accuracy = Math.round((correctCount / reviewQuestions.length) * 100);
    const sessionRecord = {
        date: new Date().toISOString().split('T')[0],
        timestamp: new Date().toISOString(),
        type: 'wrong-review',
        total: reviewQuestions.length,
        correct: correctCount,
        accuracy: accuracy,
        duration_sec: totalTime
    };
    const sessionsList = JSON.parse(localStorage.getItem('isee_sessions') || '[]');
    sessionsList.push(sessionRecord);
    localStorage.setItem('isee_sessions', JSON.stringify(sessionsList));

    // 渲染结果
    const mm = String(Math.floor(totalTime / 60)).padStart(2, '0');
    const ss = String(totalTime % 60).padStart(2, '0');

    let resultsHtml =
        '<div class="score-summary">' +
        '<div class="score-item"><span class="score-value">' + correctCount + '/' + reviewQuestions.length + '</span><span class="score-label">Score</span></div>' +
        '<div class="score-item"><span class="score-value">' + accuracy + '%</span><span class="score-label">Accuracy</span></div>' +
        '<div class="score-item"><span class="score-value">' + mm + ':' + ss + '</span><span class="score-label">Time</span></div>' +
        '</div>';

    if (wordsToRemove.length > 0) {
        resultsHtml += '<div class="mastered-list"><h3>✅ Mastered & Removed (' + wordsToRemove.length + ')</h3>';
        wordsToRemove.forEach(w => {
            resultsHtml += '<div class="mastered-item"><strong>' + w.word + '</strong> — ' + w.definition + '</div>';
        });
        resultsHtml += '</div>';
    }
    if (stillWrong.length > 0) {
        resultsHtml += '<div class="still-wrong-list"><h3>🔁 Still Need Review (' + stillWrong.length + ')</h3>';
        stillWrong.forEach(w => {
            resultsHtml += '<div class="mistake-item">' +
                '<div class="mistake-word">' + w.word + '</div>' +
                '<div class="mistake-def">' + w.definition + '</div>' +
                '</div>';
        });
        resultsHtml += '</div>';
    }

    resultsHtml += '<div class="action-links">' +
        '<a href="wrong-review.html">🔁 Review Again</a>' +
        '<a href="mistakes.html">📖 Mistakes Book</a>' +
        '<a href="history.html">📊 Progress</a>' +
        '<a href="index.html">🏠 Home</a>' +
        '</div>';

    document.getElementById('results').innerHTML = resultsHtml;
    document.getElementById('results').classList.add('show');
    document.getElementById('submit-btn').disabled = true;
    document.getElementById('submit-btn').textContent = 'SUBMITTED';
    document.getElementById('results').scrollIntoView({behavior: 'smooth'});
}

init();
</script>

</body>
</html>'''

# ============ 生成 history.html ============
def render_history_html():
    """历史成绩追踪页面 + Chart.js 折线图"""
    return '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ISEE Vocabulary — Progress History</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
* { box-sizing: border-box; }
body {
    font-family: "Times New Roman", Georgia, serif;
    max-width: 860px;
    margin: 0 auto;
    padding: 24px 20px 60px;
    background: #ffffff;
    color: #1a1a1a;
    line-height: 1.6;
}
header {
    border-bottom: 3px double #333;
    padding-bottom: 16px;
    margin-bottom: 28px;
}
h1 { font-size: 24px; margin: 0 0 6px; letter-spacing: 0.5px; }
h2 {
    font-size: 18px;
    margin: 32px 0 16px;
    padding-left: 12px;
    border-left: 4px solid #1a4d8f;
    color: #1a1a1a;
}
.subtitle { font-size: 13px; color: #555; font-style: italic; }
.back-link {
    display: inline-block;
    margin-bottom: 20px;
    color: #1a4d8f;
    text-decoration: none;
    font-family: Arial, sans-serif;
    font-size: 14px;
}
.back-link:hover { text-decoration: underline; }

/* 统计卡片 */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 14px;
    margin-bottom: 12px;
}
.stat-card {
    background: #f5f5f0;
    padding: 18px 16px;
    border: 1px solid #ddd;
    border-radius: 6px;
    text-align: center;
    font-family: "Helvetica Neue", Arial, sans-serif;
}
.stat-value {
    font-size: 28px;
    font-weight: bold;
    color: #1a4d8f;
    line-height: 1.2;
}
.stat-label {
    font-size: 11px;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 6px;
}

/* 图表区域 */
.chart-section {
    background: white;
    padding: 20px;
    border: 1px solid #ddd;
    border-radius: 6px;
    margin-top: 16px;
}
.chart-container {
    position: relative;
    height: 320px;
    margin-top: 12px;
}

/* 会话表格 */
.sessions-section { margin-top: 24px; }
table {
    width: 100%;
    border-collapse: collapse;
    background: white;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
    border-radius: 6px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
th {
    background: #1a4d8f;
    color: white;
    padding: 12px 14px;
    text-align: left;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
td {
    padding: 12px 14px;
    border-bottom: 1px solid #eee;
}
tr:last-child td { border-bottom: none; }
tr:hover td { background: #fafafa; }

.clear-section {
    margin-top: 20px;
    text-align: right;
}
button {
    background: #1a4d8f;
    color: white;
    border: none;
    padding: 8px 18px;
    font-size: 13px;
    border-radius: 4px;
    cursor: pointer;
    font-family: Arial, sans-serif;
}
button.danger { background: #c62828; }
button:hover { opacity: 0.9; }

.empty-state {
    text-align: center;
    padding: 60px 20px;
    color: #888;
}
.empty-state h2 {
    color: #555;
    font-weight: normal;
    border-left: none;
    padding-left: 0;
}

@media (max-width: 600px) {
    .stats-grid { grid-template-columns: repeat(2, 1fr); }
    .stat-value { font-size: 22px; }
    .chart-container { height: 240px; }
    table { font-size: 12px; }
    th, td { padding: 8px 6px; }
}
</style>
</head>
<body>

<a href="index.html" class="back-link">← Back to Practice</a>

<header>
    <h1>PROGRESS HISTORY</h1>
    <div class="subtitle">Track your ISEE vocabulary improvement over time</div>
</header>

<div id="content"></div>

<script>
function render() {
    const sessions = JSON.parse(localStorage.getItem('isee_sessions') || '[]');
    const content = document.getElementById('content');

    if (sessions.length === 0) {
        content.innerHTML = '<div class="empty-state">' +
            '<h2>📊 No sessions yet!</h2>' +
            '<p>Complete a practice test and your progress will appear here.</p>' +
            '</div>';
        return;
    }

    // 统计
    const totalSessions = sessions.length;
    const totalQuestions = sessions.reduce((sum, s) => sum + s.total, 0);
    const totalCorrect = sessions.reduce((sum, s) => sum + s.correct, 0);
    const avgAccuracy = Math.round(totalCorrect / totalQuestions * 100);

    // 趋势：最近 3 次 vs 最早 3 次
    let trendText = '—';
    let trendColor = '#888';
    if (sessions.length >= 4) {
        const recent = sessions.slice(-3).reduce((sum, s) => sum + s.accuracy, 0) / 3;
        const earlier = sessions.slice(0, 3).reduce((sum, s) => sum + s.accuracy, 0) / 3;
        const diff = Math.round(recent - earlier);
        if (diff > 0) { trendText = '↑ +' + diff + '%'; trendColor = '#27ae60'; }
        else if (diff < 0) { trendText = '↓ ' + diff + '%'; trendColor = '#c62828'; }
        else { trendText = '→ 0%'; }
    }

    let html = '<div class="stats-grid">' +
        '<div class="stat-card"><div class="stat-value">' + totalSessions + '</div><div class="stat-label">Sessions</div></div>' +
        '<div class="stat-card"><div class="stat-value">' + totalQuestions + '</div><div class="stat-label">Questions</div></div>' +
        '<div class="stat-card"><div class="stat-value">' + avgAccuracy + '%</div><div class="stat-label">Avg Accuracy</div></div>' +
        '<div class="stat-card"><div class="stat-value" style="color:' + trendColor + '">' + trendText + '</div><div class="stat-label">Trend</div></div>' +
        '</div>';

    html += '<h2>📈 Accuracy Over Time</h2>' +
        '<div class="chart-section"><div class="chart-container"><canvas id="accuracy-chart"></canvas></div></div>';

    html += '<h2>📋 All Sessions</h2>' +
        '<table><thead><tr>' +
        '<th>Date</th><th>Time</th><th>Score</th><th>Accuracy</th><th>Duration</th>' +
        '</tr></thead><tbody>';

    // 倒序显示（最新在上）
    sessions.slice().reverse().forEach(s => {
        const time = new Date(s.timestamp).toLocaleTimeString('en-US', {hour: '2-digit', minute: '2-digit'});
        const mins = Math.floor(s.duration_sec / 60);
        const secs = s.duration_sec % 60;
        const duration = mins + 'm ' + secs + 's';
        const accColor = s.accuracy >= 80 ? '#27ae60' : (s.accuracy >= 60 ? '#f57c00' : '#c62828');
        html += '<tr>' +
            '<td>' + s.date + '</td>' +
            '<td>' + time + '</td>' +
            '<td>' + s.correct + ' / ' + s.total + '</td>' +
            '<td style="color:' + accColor + ';font-weight:bold;">' + s.accuracy + '%</td>' +
            '<td>' + duration + '</td>' +
            '</tr>';
    });

    html += '</tbody></table>' +
        '<div class="clear-section">' +
        '<button class="danger" onclick="clearHistory()">Clear All History</button>' +
        '</div>';

    content.innerHTML = html;

    // 渲染 Chart.js 折线图
    const ctx = document.getElementById('accuracy-chart');
    const labels = sessions.map(s => s.date);
    const data = sessions.map(s => s.accuracy);

    new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Accuracy',
                data: data,
                borderColor: '#1a4d8f',
                backgroundColor: 'rgba(26, 77, 143, 0.1)',
                borderWidth: 2.5,
                pointRadius: 5,
                pointBackgroundColor: '#1a4d8f',
                pointHoverRadius: 8,
                tension: 0.3,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const idx = context.dataIndex;
                            const s = sessions[idx];
                            return [
                                'Accuracy: ' + s.accuracy + '%',
                                'Score: ' + s.correct + '/' + s.total,
                                'Duration: ' + Math.floor(s.duration_sec / 60) + 'm ' + (s.duration_sec % 60) + 's'
                            ];
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    ticks: {
                        callback: function(value) { return value + '%'; }
                    },
                    title: { display: true, text: 'Accuracy (%)' }
                },
                x: {
                    ticks: { maxRotation: 45, minRotation: 0 }
                }
            }
        }
    });
}

function clearHistory() {
    if (confirm('Clear all session history? This cannot be undone.')) {
        localStorage.removeItem('isee_sessions');
        render();
    }
}

render();
</script>

</body>
</html>'''

# ============ 生成 mistakes.html ============
def render_mistakes_html():
    """错题本页面（从 localStorage 读数据）"""
    return '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ISEE Vocabulary — Mistakes Book</title>
<style>
* { box-sizing: border-box; }
body {
    font-family: "Times New Roman", Georgia, serif;
    max-width: 760px;
    margin: 0 auto;
    padding: 24px 20px 60px;
    background: #ffffff;
    color: #1a1a1a;
    line-height: 1.6;
}
header {
    border-bottom: 3px double #333;
    padding-bottom: 16px;
    margin-bottom: 28px;
}
h1 { font-size: 24px; margin: 0 0 6px; letter-spacing: 0.5px; }
.subtitle { font-size: 13px; color: #555; font-style: italic; }
.stats {
    background: #f5f5f0;
    padding: 14px 18px;
    border: 1px solid #ddd;
    border-radius: 4px;
    margin-bottom: 24px;
    font-family: "Helvetica Neue", Arial, sans-serif;
    font-size: 14px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.mistake-card {
    background: white;
    padding: 18px 20px;
    margin-bottom: 14px;
    border-left: 4px solid #f44336;
    border-radius: 0 4px 4px 0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
.mistake-word {
    font-size: 20px;
    font-weight: bold;
    color: #1a1a1a;
    margin-bottom: 6px;
}
.mistake-type {
    display: inline-block;
    font-size: 11px;
    background: #1a4d8f;
    color: white;
    padding: 2px 8px;
    border-radius: 3px;
    margin-left: 8px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    vertical-align: middle;
}
.mistake-def {
    color: #333;
    font-size: 15px;
    font-style: italic;
    margin-bottom: 12px;
    padding: 8px 12px;
    background: #f9f9f7;
    border-radius: 4px;
}
.mistake-stem {
    font-size: 14px;
    color: #444;
    margin-bottom: 8px;
}
.mistake-options { font-size: 14px; margin-top: 8px; }
.opt-row { margin: 3px 0; padding: 4px 8px; border-radius: 3px; }
.opt-correct { background: #e8f5e9; color: #1e7e3e; font-weight: 600; }
.opt-wrong { background: #ffebee; color: #c62828; text-decoration: line-through; }
.opt-neutral { color: #666; }
.mistake-date {
    font-size: 11px;
    color: #999;
    margin-top: 10px;
    text-align: right;
    font-family: Arial, sans-serif;
}
button {
    background: #1a4d8f;
    color: white;
    border: none;
    padding: 8px 18px;
    font-size: 13px;
    border-radius: 4px;
    cursor: pointer;
    font-family: Arial, sans-serif;
}
button.danger { background: #c62828; }
button:hover { opacity: 0.9; }
.empty-state {
    text-align: center;
    padding: 60px 20px;
    color: #888;
}
.empty-state h2 { color: #555; font-weight: normal; }
.back-link {
    display: inline-block;
    margin-bottom: 20px;
    color: #1a4d8f;
    text-decoration: none;
    font-family: Arial, sans-serif;
    font-size: 14px;
}
.back-link:hover { text-decoration: underline; }
</style>
</head>
<body>

<a href="index.html" class="back-link">← Back to Practice</a>

<header>
    <h1>MISTAKES BOOK</h1>
    <div class="subtitle">Words to review for ISEE Lower Level vocabulary</div>
</header>

<div id="content">
    <!-- populated by JS -->
</div>

<script>
function render() {
    const mistakes = JSON.parse(localStorage.getItem('isee_mistakes') || '[]');
    const content = document.getElementById('content');

    if (mistakes.length === 0) {
        content.innerHTML = '<div class="empty-state">' +
            '<h2>📚 No mistakes yet!</h2>' +
            '<p>Take the practice test and your mistakes will appear here for review.</p>' +
            '</div>';
        return;
    }

    // 按日期倒序
    mistakes.sort((a, b) => (b.date || '').localeCompare(a.date || ''));

    let html = '<div class="stats">' +
        '<span><strong>' + mistakes.length + '</strong> word(s) to review</span>' +
        '<button class="danger" onclick="clearMistakes()">Clear All</button>' +
        '</div>';

    mistakes.forEach(m => {
        const typeLabel = m.type === 'synonym' ? 'Synonym' : 'Sentence';
        let optsHtml = '<div class="mistake-options">';
        m.options.forEach((opt, idx) => {
            let cls = 'opt-neutral';
            let prefix = '   ';
            if (idx === m.correct_index) {
                cls = 'opt-correct';
                prefix = '✓ ';
            } else if (idx === m.user_index) {
                cls = 'opt-wrong';
                prefix = '✗ ';
            }
            const letter = String.fromCharCode(65 + idx);
            optsHtml += '<div class="opt-row ' + cls + '">' + prefix + letter + '. ' + opt + '</div>';
        });
        optsHtml += '</div>';

        html += '<div class="mistake-card">' +
            '<div class="mistake-word">' + m.word + '<span class="mistake-type">' + typeLabel + '</span></div>' +
            '<div class="mistake-def">' + m.definition + '</div>' +
            '<div class="mistake-stem">' + m.stem + '</div>' +
            optsHtml +
            '<div class="mistake-date">' + (m.date || 'unknown') + '</div>' +
            '</div>';
    });

    content.innerHTML = html;
}

function clearMistakes() {
    if (confirm('Clear all mistake records? This cannot be undone.')) {
        localStorage.removeItem('isee_mistakes');
        render();
    }
}

render();
</script>

</body>
</html>'''

# ============ 写入文件 ============
with open(INDEX_FILE, "w", encoding="utf-8") as f:
    f.write(render_index_html(questions, timestamp))

with open(MISTAKES_FILE, "w", encoding="utf-8") as f:
    f.write(render_mistakes_html())
with open(HISTORY_FILE, "w", encoding="utf-8") as f:
    f.write(render_history_html())
with open(WRONG_REVIEW_FILE, "w", encoding="utf-8") as f:
    f.write(render_wrong_review_html())

print(f"📄 已生成：")
print(f"   - {INDEX_FILE} （做题入口）")
print(f"   - {MISTAKES_FILE} （错题本）")
print(f"   - {HISTORY_FILE} （历史成绩 + 折线图）")
print(f"   - {WRONG_REVIEW_FILE} （错题集中复习）")
print(f"\n✅ 全部完成！本地双击 index.html 测试")
print(f"📊 词库利用率：{TOTAL_QUESTIONS}/{len(word_pool)} = {TOTAL_QUESTIONS*100//len(word_pool)}%")