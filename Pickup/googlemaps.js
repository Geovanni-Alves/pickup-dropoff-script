
const md5 = (key = '') => {
  const code = key.toLowerCase().replace(/\s/g, '');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, code)
    .map((char) => (char + 256).toString(16).slice(-2))
    .join('');
};

const getCache = (key) => {
  return CacheService.getDocumentCache().get(md5(key));
};

// Store the results for 6 hours
const setCache = (key, value) => {
  const expirationInSeconds = 6 * 60 * 60;
  CacheService.getDocumentCache().put(md5(key), value, expirationInSeconds);
};

function test_google  () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actualSheetName = ss.getActiveSheet().getName();
  var sheet = ss.getSheetByName(actualSheetName);

  var origin = '2440 Main St, Vancouver, BC V5T 3E2';
  var destination = '480 Robson St #707, Vancouver, BC V6B 1S1';
  var time = sheet.getRange("H3").getValue();
  console.log(time)
  //console.log('distance: ',GOOGLEMAPS_DISTANCE(origin,destination,'driving'));
  console.log('distance: ',GOOGLEMAPS_DURATION(origin,destination));

}

// const GOOGLEMAPS_DURATION = (origin, destination, mode = 'driving') => {
//   // Generate a unique cache key for this request
//   //const departureTimestamp = new Date(`1970-01-01T${departureTime}`).getTime() / 1000;
//   const cacheKey = ['duration', origin, destination, mode].join(',');
  

//   // Check if the result is in the cache
//   const cachedResult = getCache(cacheKey);
//   if (cachedResult !== null) {
//     return cachedResult;
//   }

//   //const apiKeyParam = `key=${GOOGLE_MAPS_API_KEY}`;
//   const apiUrl = `https://maps.googleapis.com/maps/api/directions/json?origin=${encodeURIComponent(origin)}&destination=${encodeURIComponent(destination)}&mode=${mode}&departure_time=now`;
  
//   // &departure_time=${departureTimestamp}`;
//   const response = UrlFetchApp.fetch(apiUrl);
//   //console.log(response.getContentText());
//   const data = JSON.parse(response.getContentText());

//   if (data.status !== 'OK' || data.routes.length === 0) {
//     throw new Error('No route found!');
//   }

//   const { legs: [{ duration: { text: time } } = {}] } = data.routes[0];
//   // Extract the numerical part of the duration and convert it to a number
//   const numericalTime = parseFloat(time.replace(' mins', '').replace('min',''));
//   setCache(cacheKey, numericalTime);
//   return numericalTime;
// };


// const GOOGLEMAPS_DISTANCE = (origin, destination, mode = 'driving') => {
//   const cacheKey = ['distance', origin, destination, mode].join(',');

//   // Check if the result is in the cache
//   const cachedResult = getCache(cacheKey);
//   if (cachedResult !== null) {
//     return cachedResult;
//   }

//   //const apiKeyParam = `key=${GOOGLE_MAPS_API_KEY}`;
//   const apiUrl = `https://maps.googleapis.com/maps/api/directions/json?origin=${encodeURIComponent(origin)}&destination=${encodeURIComponent(destination)}&mode=${mode}`;
  

//   const response = UrlFetchApp.fetch(apiUrl);
//   //console.log(response.getContentText());
//   const data = JSON.parse(response.getContentText());

//   if (data.status !== 'OK' || data.routes.length === 0) {
//     throw new Error('No route found!');
//   }
//  const { legs: [{ distance: { text: distance } } = {}] } = data.routes[0];

//   // Cache the result for future use
//   setCache(cacheKey, distance);
//   return distance

// };








/* Calculate the distance between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DISTANCE("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The distance in miles
 * @customFunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode) => {
  const key = ['duration', origin, destination, mode].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();

  if (!data) {
    throw new Error('No route found!');
  }

  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
    setCache(distance);
  return distance;
};

/**
 * Use Reverse Geocoding to get the address of
 * a point location (latitude, longitude) on Google Maps.
 *
 * =GOOGLEMAPS_REVERSEGEOCODE(latitude, longitude)
 *
 * @param {String} latitude The latitude to lookup.
 * @param {String} longitude The longitude to lookup.
 * @return {String} The postal address of the point.
 * @customFunction
 */

const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(latitude, longitude);
  return data.formatted_address;
};

/**
 * Get the latitude and longitude of any
 * address on Google Maps.
 *
 * =GOOGLEMAPS_LATLONG("10 Hanover Square, NY")
 *
 * @param {String} address The address to lookup.
 * @return {String} The latitude and longitude of the address.
 * @customFunction
 */
const GOOGLEMAPS_LATLONG = (address) => {
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error('Address not found!');
  }
  const { geometry: { location: { lat, lng } } = {} } = data;
  return `${lat}, ${lng}`;
};

/**
 * Find the driving direction between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DIRECTIONS("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The driving direction
 * @customFunction
 */
const GOOGLEMAPS_DIRECTIONS = (origin, destination, mode = 'driving') => {
  const { routes = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!routes.length) {
    throw new Error('No route found!');
  }
  return routes
    .map(({ legs }) => {
      return legs.map(({ steps }) => {
        return steps.map((step) => {
          return step.html_instructions.replace(/<[^>]+>/g, '');
        });
      });
    })
    .join(', ');
};



/**
 * Calculate the travel time between two locations
 * on Google Maps.
 *
 * =GOOGLEMAPS_DURATION("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The time in minutes
 * @customFunction
 */
const GOOGLEMAPS_DURATION = (origin, destination, mode = 'driving') => {
  const key = ['duration', origin, destination, mode].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!data) {
    throw new Error('No route found!');
  }
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  // Store the result in internal cache for future
   const numericalTime = parseFloat(time.replace(' mins', '').replace('min',''));
  setCache(key, numericalTime);
  return numericalTime;
};


const geoCodeAddress = (address) => {
  const key = [address].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  var response = Maps.newGeocoder().geocode(address);
  
  if (response.status === 'OK') {
    var location = response.results[0].geometry.location;
    var lat = location.lat;
    var lng = location.lng;
    var latLng = lat + ", " + lng;
    setCache(latLng);
    return latLng
  } else {
    return "Error";
  }
}






