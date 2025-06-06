function parseSynonyms(xmlString) {
  let parser = new DOMParser();
  let xmlDoc = parser.parseFromString(xmlString, 'text/xml');
  let synonymDict = {};

  let entries = xmlDoc.getElementsByTagName('SynmDictionary');
  for (let entry of entries) {
    let type = entry.getElementsByTagName('Type')[0].textContent.trim();
    let synonyms = entry.getElementsByTagName('Synm')[0].textContent.trim().split('|').filter(Boolean);
    let word = entry.getElementsByTagName('Word')[0].textContent.trim();

    synonyms.forEach(synonym => {
            synonymDict[synonym.toUpperCase()] = word;
    });
  }
  return synonymDict;
}
