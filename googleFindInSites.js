/**
 * this function will create a link to search a term on a number of websites with google
 * https://www.google.com/search?q=button+site%3Amaterial-ui.com+OR+site%3Aant.design+OR+site%3Aschema.org
 * @param {array of strings} sites the sites the search is limited to
 * @param {string} searchTerm the term that is being searched
 * @returns 
 */
function googleFindInSites(sites, searchTerm){
	let googleBase = "https://www.google.com/search?q=";
	let siteStart = "site%3A";
	let siteOR = "OR";
	let separator = "+";
	let rvArr = [googleBase,searchTerm];
	let sitesMapped = sites.map(site => {
		return `${siteStart}${site}`;
	});
	rvArr.push(sitesMapped.reduce((prev, cur, curIdx) => {
		if(curIdx == 0 ) return prev+separator+cur;
		return prev + separator + siteOR + separator + cur;
	}));
	return rvArr.reduce((prev, cur, curIdx) => {
		if(curIdx == 1) return prev+cur;
		return prev + separator + cur;
	});
}

let term = "button";
googleFindInSites(["material-ui.com","ant.design","schema.org"], term)