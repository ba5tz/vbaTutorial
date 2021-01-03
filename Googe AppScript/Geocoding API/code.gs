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

function hitungJarak(asal,tujuan){
  
  var mapobj = Maps.newDirectionFinder();
  mapobj.setOrigin(asal);
  mapobj.setDestination(tujuan);
  var hasil = mapobj.getDirections();

  return hasil.routes[0].legs[0].distance.value;
}

function hitungJarakLat(lat1,long1,lat2,long2){
  
  var mapobj = Maps.newDirectionFinder();
  mapobj.setOrigin(lat1, long1)
  mapobj.setDestination(lat2,long2);
  var hasil = mapobj.getDirections();

  return hasil.routes[0].legs[0].distance.value;

}
