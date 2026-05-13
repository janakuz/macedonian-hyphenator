/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// @ts-check
/** @type {Record<string, number>} */
const values = {
'а': 12,
'б': 2,
'в': 2,
'г': 2,
'д': 2,
'ѓ': 2,
'е': 12,
'ж': 2,
'з': 2,
'ѕ': 2,
'и': 12,
'ј': 3,
'к': 1,
'л': 3,
'љ': 3,
'м': 3,
'н': 3,
'њ': 3,
'о': 12,
'п': 2,
'р': 6,
'с': 1,
'т': 1,
'ќ': 1,
'у': 12,
'ф': 1,
'х': 1,
'ц': 1,
'ч': 1,
'џ': 2,
'ш': 1,
'FC': 0,
'`': 0,
'S': 0,
'E': 0
}


function is_consonant(letter){
    return "бвгдѓжзѕјклљмнњпрстќфхцчџш".includes(letter.toLowerCase())
}

function is_vowel(letter){
    return "аеиоу".includes(letter.toLowerCase())
}

function prepare(word){
    let new_word = 'S' + word + 'E'
    let j = 1
    for (let i = 0; i < word.length-1; i++){
        if (is_vowel(word[i]) && is_vowel(word[i+1])){
            new_word = new_word.slice(0,i+j+1) + "FC" + new_word.slice(i+j+1)
            j += 2
        }
    }
    return new_word
}

function triplet_diff(a, b , c, values){
    return values[a] - values[b] - values[c]
}


function get_char_arr(word = ""){
    let chars = []
    for (let c of word){
        if (!"FC".includes(c))
            chars.push(c)
        else if (c == "F"){
            chars.push("FC")
        }
    }
    return chars
}

function find_nuclei(word, values, special){
  /** @type {[string, number][]} */
    let nuclei = []
    let chars = get_char_arr(word)
    for (let i = 1; i < chars.length-1; i++){
        if ((triplet_diff(chars[i], chars[i-1], chars[i+1], values) > 0 && (is_vowel(chars[i]) || special.includes(chars[i]))) ||
            (triplet_diff(chars[i], chars[i-1], chars[i+1], values) == 0 && is_vowel(chars[i])) ||
    (special.includes(chars[i]) && i != 0 && i != word.length-1 && is_consonant(chars[i-1]) && is_consonant(chars[i+1])))
            nuclei.push([chars[i], i])
    }
    return nuclei
}



function split_syllables_mk(word){
    let new_word = word.toLowerCase()
    new_word = prepare(new_word)
    let nuclei = find_nuclei(new_word, values, "р")
    let chars = get_char_arr(new_word)

    /** @type {string[]} */
    let syllables = []

    if (nuclei.length == 0)
        syllables = [word.toLowerCase()]

    let i = 1
    let j = 0
    while (i < chars.length && j < nuclei.length){
        let curr_nucl = nuclei[j]
        let syllable = ""
        let nucl_added = false
        if (!nucl_added){
            while (i < chars.length && i < curr_nucl[1]){
                if (chars[i] != "FC")
                    syllable += chars[i]
                i += 1
            }
            nucl_added = true
        }
        if (i < chars.length && chars[i] == "FC"){
            syllables.push(syllable)
            i += 1
        }
        else{
            while (i < chars.length-1 && values[chars[i]] > values[chars[i+1]]){
                syllable += chars[i]
                i += 1
            }
            if (is_consonant(chars[i]) && is_consonant(chars[i+1]) && !nuclei.includes(['р', i+1])){
                syllable += chars[i]
                i += 1
            }
            if (["ство", "ства","штво", "штва"].includes(chars.slice(i-1,i+3).join('')) ||
                chars.slice(i-1,i+4).join('') == "ствен"){
                syllable = syllable.slice(0,-1)
                i -= 1
            }
            if (["ски","ска","ско"].includes(chars.slice(i-1,i+2).join(''))){
                syllable = syllable.slice(0,-1)
                i -= 1
            }
            if (j == nuclei.length-1 && syllables.length == nuclei.length-1){
                syllable += chars.slice(i,-1).join('')
                i = chars.length
            }
            syllables.push(syllable)
        }
        j += 1
    }
    return syllables
}


function filterMacedonianSyllables(syllables) {
  if (syllables.length <= 1) return syllables;

  if (syllables[0].length === 1) {
    syllables[1] = syllables[0] + syllables[1];
    syllables.shift();
  }

  if (syllables.length > 1 && syllables[syllables.length - 1].length === 1) {
    syllables[syllables.length - 2] = syllables[syllables.length - 2] + syllables[syllables.length - 1];
    syllables.pop();
  }

  return syllables;
}



function getHyphenatedWord(word) {
  let rawSyllables = split_syllables_mk(word.toLowerCase());
  let legalSyllables = filterMacedonianSyllables(rawSyllables);
  let hyphenatedLower = legalSyllables.join('\u001f');  
  
  let result = "";
  let originalIndex = 0;
  
  for (let i = 0; i < hyphenatedLower.length; i++) {
    if (hyphenatedLower[i] === '\u001f') {
      result += '\u001f';
    } else {
      result += word[originalIndex];
      originalIndex++;
    }
  }
  return result;
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    paragraphs.items.forEach((para) => {
      let originalText = para.text;
      let cleanText = originalText.replace(/\u001f/g, "");

      let words = cleanText.split(" ");
      let processedWords = words.map(w => getHyphenatedWord(w));      
      para.insertText(processedWords.join(" "), Word.InsertLocation.replace);
    });


    await context.sync();
  });
}

console.log(getHyphenatedWord("Македонија")); 
