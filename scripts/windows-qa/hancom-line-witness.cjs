'use strict';

/** Locate the two literal anchors independently in each saved file. Their
 * positions in the XML change when the editor inserts page carrier paragraphs. */
module.exports = function hancomLineWitness(original, saved) {
  const normalized = value => value.replace(/\s/gu, '');
  const authored = normalized(original.getTextRange(0, 127, 0, 0xffffffff));
  const next = normalized(original.getTextRange(0, 135, 0, 0xffffffff));
  if (!authored || !next) throw Error('The pinned line anchors are missing');
  const locate = literal => {
    const matches = [];
    for (let index = 0; index < saved.getParagraphCount(0); index++) {
      if (normalized(saved.getTextRange(0, index, 0, 0xffffffff)) === literal) matches.push(index);
    }
    if (matches.length !== 1) throw Error('A line anchor was missing or ambiguous');
    return matches[0];
  };
  const paragraph = locate(authored);
  const page = JSON.parse(saved.getPageOfPosition(0, paragraph)).page;
  const nextPage = JSON.parse(saved.getPageOfPosition(0, locate(next))).page;
  const runs = JSON.parse(saved.getPageTextLayout(page)).runs.filter(run =>
    run.secIdx === 0 && run.paraIdx === paragraph && run.parentParaIdx === undefined && normalized(run.text));
  if (normalized(runs.map(run => run.text).join('')) !== authored) throw Error('Line geometry did not cover its literal text');
  return { page: page + 1, nextPage: nextPage + 1, lines: new Set(runs.map(run => run.y)).size };
};
