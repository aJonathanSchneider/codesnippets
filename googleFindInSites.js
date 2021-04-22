function googleFindInSites(sites, searchTerm){
	let rv="";
	let googleBase = "https://www.google.com/search?q=";
	let siteStart = "site%3A";
	let siteOR = "OR";
	let separator = "+";
	//site%3Amaterial-ui.com+OR+site%3Aant.design+button
	let rvArr = [googleBase+searchTerm];
	let sitesMapped = sites.map(site => {
		return `${siteStart}${site}`;
	});
	rvArr.push(sitesMapped.reduce((prev, cur, curIdx) => {
		if(curIdx == 0 || curIdx == sitesMapped.length - 1) return prev+cur;
		return prev + separator + siteOR + separator + cur;
	},""));
	return rvArr.reduce((prev, cur, curIdx) => {
		if(curIdx == 0 || curIdx == rvArr.length - 1) return prev+cur;
		return prev + separator + cur;
	})
}

let term = "button";
googleFindInSites(["material-ui.com","ant.design","schema.org"], term)