// === script.js (duplicate triangle + dictionary lookup) ===

const sheetURL = 'https://raw.githubusercontent.com/RJ-Flashcards/Flashcard-app3/main/vocab.csv';

let flashcards = [];
let currentCard = 0;
let isFlipped = false;

/* ---------------------------
   Small CSV parser (handles quotes & commas)
---------------------------- */
function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        cell += '"'; // escaped quote ""
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        row.push(cell);
        cell = '';
      } else if (ch === '\n') {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = '';
      } else if (ch === '\r') {
        // ignore CR in CRLF
      } else {
        cell += ch;
      }
    }
  }
  // flush last cell/row
  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  // trim whitespace around each cell
  return rows.map(r => r.map(c => (c ?? '').trim()));
}

/* ---------------------------------------------
   Duplicate detection (case-insensitive)
---------------------------------------------- */
function buildTermCounts(cards) {
  const counts = {};
  cards.forEach(({ term }) => {
    const key = (term || '').trim().toLowerCase();
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function tagDuplicates(cards) {
  const counts = buildTermCounts(cards);
  return cards.map(c => ({
    ...c,
    isDuplicate: counts[(c.term || '').trim().toLowerCase()] > 1
  }));
}

/* ---------------------------
   Data loading & preparation
---------------------------- */
function fetchFlashcards() {
  fetch(sheetURL)
    .then(response => {
      if (!response.ok) throw new Error('Failed to fetch CSV');
      return response.text();
    })
    .then(text => {
      // Parse CSV robustly
      let rows = parseCSV(text);

      // Auto-detect and drop header if it looks like one
      if (rows.length && rows[0].length >= 2) {
        const h0 = rows[0][0].toLowerCase();
        const h1 = rows[0][1].toLowerCase();
        const looksLikeHeader =
          (h0.includes('word') || h0.includes('term')) &&
          (h1.includes('def') || h1.includes('meaning'));
        if (looksLikeHeader) rows = rows.slice(1);
      }

      // Map to {term, definition}, ignore empty terms
      flashcards = rows
        .filter(r => r && r.length >= 1 && (r[0] ?? '').trim() !== '')
        .map(r => ({
          term: (r[0] ?? '').trim(),
          definition: (r[1] ?? '').trim(),
        }));

      // Tag duplicates before shuffling
      flashcards = tagDuplicates(flashcards);

      shuffleFlashcards();
      displayCard();
    })
    .catch(err => {
      document.getElementById('card-front').innerText = 'Error loading flashcards.';
      console.error('Error:', err);
    });
}

function shuffleFlashcards() {
  for (let i = flashcards.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [flashcards[i], flashcards[j]] = [flashcards[j], flashcards[i]];
  }
}

/* ---------------------------
   Rendering
---------------------------- */
function displayCard() {
  const front = document.getElementById('card-front');
  const back = document.getElementById('card-back');
  const card = flashcards[currentCard];

  // Clear and render front
  front.textContent = ''; // reset entirely before adding nodes

  // Word node
  const wordNode = document.createElement('span');
  wordNode.textContent = card.term || '';
  front.appendChild(wordNode);

  // Triangle icon if duplicate (▲)
  if (card.isDuplicate) {
    const icon = document.createElement('span');
    icon.className = 'dup-flag';
    icon.title = 'Duplicate word';
    icon.setAttribute('aria-label', 'Duplicate word');
    icon.textContent = '▲';
    icon.style.marginLeft = '0.4rem';
    icon.style.opacity = '0.9';
    front.appendChild(icon);
  }

  // Back
  back.textContent = card.definition || '';

  // Keep current flipped state
  const flashcardEl = document.getElementById('flashcard');
  if (isFlipped) flashcardEl.classList.add('flipped');
  else flashcardEl.classList.remove('flipped');

  // 👇 expose current word for dictionary lookup
  window.currentWord = card.term || '';
}

/* ---------------------------
   Interaction
---------------------------- */
// Flip only on card tap (never on button press)
document.getElementById('flashcard').addEventListener('click', (e) => {
  if (e.target.tagName.toLowerCase() === 'button') {
    e.stopPropagation();
    return;
  }
  const flashcard = document.getElementById('flashcard');
  flashcard.classList.toggle('flipped');
  isFlipped = !isFlipped;
});

// Next / Back, preserve flip state
document.getElementById('next-btn')?.addEventListener('click', (e) => {
  e.stopPropagation();
  currentCard = (currentCard + 1) % flashcards.length;
  displayCard();
});

document.getElementById('back-btn')?.addEventListener('click', (e) => {
  e.stopPropagation();
  currentCard = (currentCard - 1 + flashcards.length) % flashcards.length;
  displayCard();
});

/* ---------------------------
   Dictionary lookup feature
---------------------------- */
// Direct-entry URL templates
const DICT_DIRECT = {
  cambridge: (w) =>
    `https://dictionary.cambridge.org/dictionary/learner-english/${encodeURIComponent(
      w.replace(/\s+/g, "-").toLowerCase()
    )}`,
  oxford: (w) =>
    `https://www.oxfordlearnersdictionaries.com/definition/english/${encodeURIComponent(
      w.replace(/\s+/g, "-").toLowerCase()
    )}`,
  longman: (w) =>
    `https://www.ldoceonline.com/dictionary/${encodeURIComponent(
      w.replace(/\s+/g, "-").toLowerCase()
    )}`,
};

// Search fallback URLs
const DICT_SEARCH = {
  cambridge: (w) =>
    `https://www.google.com/search?q=${encodeURIComponent(
      `site:dictionary.cambridge.org "${w}"`
    )}`,
  oxford: (w) =>
    `https://www.google.com/search?q=${encodeURIComponent(
      `site:oxfordlearnersdictionaries.com "${w}"`
    )}`,
  longman: (w) =>
    `https://www.google.com/search?q=${encodeURIComponent(
      `site:ldoceonline.com "${w}"`
    )}`,
};

// Heuristic: detect phrases / inflections
function shouldUseSearch(word) {
  if (!word) return true;
  const w = word.trim();

  if (/\s/.test(w)) return true; // multi-word
  if (/[^\p{L}\p{M}-]/u.test(w)) return true; // punctuation

  const lw = w.toLowerCase();
  if (/(ing|ed|ies|ied|ers|est|s)$/.test(lw) && !/(ss|us|is|as)s$/.test(lw)) {
    return true;
  }

  return false;
}

function buildDictUrl(dictKey, word) {
  const clean = (word || "").trim();
  if (shouldUseSearch(clean)) {
    return DICT_SEARCH[dictKey](clean);
  }
  return DICT_DIRECT[dictKey](clean);
}

// Initialize dropdown + button after DOM is ready
document.addEventListener("DOMContentLoaded", () => {
  const helper = document.getElementById("dictHelper");
  const choice = document.getElementById("dictChoice");
  const btn = document.getElementById("lookupBtn");

  if (!helper || !choice || !btn) return;

  helper.addEventListener("click", (e) => e.stopPropagation());
  helper.addEventListener("touchstart", (e) => e.stopPropagation(), { passive: true });

  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const dictKey = choice.value;

    // Prefer window.currentWord; fall back to the flashcard array
    let word = (window.currentWord || "").trim();
    if (!word && Array.isArray(flashcards) && flashcards.length > 0) {
      const card = flashcards[currentCard] || {};
      word = (card.term || "").trim();
    }

    if (!word) {
      alert("No word is selected yet.");
      return;
    }

    const url = buildDictUrl(dictKey, word);
    window.open(url, "_blank", "noopener");
  });
});

// Go!
fetchFlashcards();
