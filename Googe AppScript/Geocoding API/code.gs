function ubahgeocode(alamat){
  var jsn = Maps.newGeocoder().geocode(alamat);
  for (var i=0;i< jsn.results.length;i++){
       var res = jsn.results[i];
    return res.geometry.location.lat + ", " + res.geometry.location.lng;
       }

}

function reversegeocode(lat, long){
  var jsn = Maps.newGeocoder().reverseGeocode(lat, long);
  return jsn.results[0].formatted_address;

}
