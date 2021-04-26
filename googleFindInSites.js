/**
 * this function will create a link to search a term on a number of websites with google
 * https://www.google.com/search?q=button+site%3Amaterial-ui.com+OR+site%3Aant.design+OR+site%3Aschema.org
 * @param {array of strings} sites the sites the search is limited to
 * @param {string} searchTerm the term that is being searched
 * @returns {string} a link for a search query
 */
function googleFindInSites(sites, searchTerm) {
  let googleBase = "https://www.google.com/search?q=";
  let siteStart = "site%3A";
  let siteOR = "OR";
  let separator = "+";
  let rvArr = [googleBase, encodeURIComponent(searchTerm)];
  let sitesMapped = sites.map((site) => {
    return `${siteStart}${site}`;
  });
  rvArr.push(
    sitesMapped.reduce((prev, cur, curIdx) => {
      if (curIdx == 0) return prev + separator + cur;
      return prev + separator + siteOR + separator + cur;
    })
  );
  return rvArr.reduce((prev, cur, curIdx) => {
    if (curIdx == 1) return prev + cur;
    return prev + separator + cur;
  });
}

function googleFindManyTermsInSites(sites, searchTerms) {
  return searchTerms
    .map((val) => googleFindInSites(sites, val))
    .reduce((prev, cur) => prev + "\n" + cur, "");
}

//enter your search terms here between the quotation marks "", and separate each term with a comma
let mySearchTerms = [
  "Brad Pitt",
  "Angelina Jolie",
  "Elon Musk",
  "Nomadland",
  "Parasite",
];

//enter the websites you want to do those searches in
let limitMySearchToSites = ["imdb.com", "de.wikipedia.org", "twitter.com"];

console.log(googleFindManyTermsInSites(limitMySearchToSites, mySearchTerms));
