function траты(inputrange) {
  let tmparr = inputrange.map(e => parseFloat(e));
  let sum = tmparr.reduce((a,b) => a+b,0);
  tmparr = tmparr.map(e => (e*10 < sum)? 0: e);
  let tmparr2 = tmparr.filter(e => e*10 >= sum);
  let minvolume = Math.min(...tmparr2);
  let dist = inputrange.map(e => Math.round(parseFloat(e)/minvolume)).join(" - ");
 
  return dist;
}