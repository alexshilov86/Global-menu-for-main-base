function get_vol(vol_list) {
  vol_ans = 0;

  try {
    for (vol of vol_list) {
      vol = vol.replaceAll(',', '.')
      Logger.log(vol);
      if (vol.toLowerCase().includes('паллет')){
        // const matches = vol.match(/([0-9]{1,}.?[0-9]{1,})[ ]{0,}паллет/i);
        const matches = vol.match(/(-?\d+[,.?\d+]*)[ ]{0,}паллет/i);
        vol_ans = vol_ans + Number(matches[1]);
        
      }
    }  
  } catch (e) {
    Logger.log(e);
    return 0
    
  }
  return vol_ans;
}

function test_get_vol() {
  art = ['НЕСУЩЕСТВЕННО', '5 паллет', ' 2 паллета']
  Logger.log(get_vol(art)+3)
  
}

