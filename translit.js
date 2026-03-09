const translitMap = {
  а: 'a', б: 'b', в: 'v', г: 'g', д: 'd', е: 'e', ё: 'yo', ж: 'j', з: 'z',
  и: 'i', й: 'y', к: 'k', л: 'l', м: 'm', н: 'n', о: 'o', п: 'p', р: 'r',
  с: 's', т: 't', у: 'u', ф: 'f', х: 'x', ц: 's', ч: 'ch', ш: 'sh', щ: 'sh',
  ы: 'y', э: 'e', ю: 'yu', я: 'ya', ь: '', ъ: '',
  қ: 'q', ғ: "g'", ў: "o'", ҳ: 'h'
};

function transliterate(text) {
  let result = '';

  for (const char of text) {
    const lower = char.toLowerCase();

    if (!(lower in translitMap)) {
      result += char;
      continue;
    }

    const translit = translitMap[lower];

    if (char === lower) {
      result += translit;
    } else {
      result += translit.charAt(0).toUpperCase() + translit.slice(1);
    }
  }

  return result;
}

module.exports = { transliterate };
