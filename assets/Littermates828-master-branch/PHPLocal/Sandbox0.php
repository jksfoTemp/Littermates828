<!DOCTYPE html>
<html>
   <head>
      <title>Sandbox0.php</title>
   </head>
   <body>
      <p>Hello World</p>
      <P />&nbsp; 
      
      <?PHP
            /* foreach($_SERVER as $key_name => $key_value) {
                print $key_name." = ".$key_value."<br>";
            } */
            print "PHP"; 
            
            ?>

            <?php 
            /*

            // prints: mysql link
            $c = mysql_connect();
            echo get_resource_type($c) . "\n";

            // prints: stream
            $fp = fopen("foo", "w");
            echo get_resource_type($fp) . "\n";

            // prints: domxml document
            $doc = new_xmldoc("1.0");
            echo get_resource_type($doc->doc) . "\n";
            */
            ?>

     <P />&nbsp; 
     <form name="myForm"  method="get" action="Sandbox0.php">
      <input type="search" name="search" /><br />
      <input type="submit" name="submit" value="Search" /><br />
  </form>
  <P />&nbsp; 
  
  <script type="text/javascript"> 
  alert("sfs");
  // https://stackoverflow.com/questions/8180296/what-information-can-we-access-from-the-client
var info={

    timeOpened:new Date(),
    timezone:(new Date()).getTimezoneOffset()/60,

    pageon(){return window.location.pathname},
    referrer(){return document.referrer},
    previousSites(){return history.length},

    browserName(){return navigator.appName},
    browserEngine(){return navigator.product},
    browserVersion1a(){return navigator.appVersion},
    browserVersion1b(){return navigator.userAgent},
    browserLanguage(){return navigator.language},
    browserOnline(){return navigator.onLine},
    browserPlatform(){return navigator.platform},
    javaEnabled(){return navigator.javaEnabled()},
    dataCookiesEnabled(){return navigator.cookieEnabled},
    dataCookies1(){return document.cookie},
    dataCookies2(){return decodeURIComponent(document.cookie.split(";"))},
    dataStorage(){return localStorage},

    sizeScreenW(){return screen.width},
    sizeScreenH(){return screen.height},
    sizeDocW(){return document.width},
    sizeDocH(){return document.height},
    sizeInW(){return innerWidth},
    sizeInH(){return innerHeight},
    sizeAvailW(){return screen.availWidth},
    sizeAvailH(){return screen.availHeight},
    scrColorDepth(){return screen.colorDepth},
    scrPixelDepth(){return screen.pixelDepth},


    latitude(){return position.coords.latitude},
    longitude(){return position.coords.longitude},
    accuracy(){return position.coords.accuracy},
    altitude(){return position.coords.altitude},
    altitudeAccuracy(){return position.coords.altitudeAccuracy},

    heading(){return position.coords.heading},
    speed(){return position.coords.speed},
    timestamp(){return position.timestamp},


 

   };

   document.getElementById("myForm").addEventListener("submit", function (e) {
  e.preventDefault();

  var formData = new FormData(form);
  // output as an object
  console.log(Object.fromEntries(formData));

  // ...or iterate through the name-value pairs
  for (var pair of formData.entries()) {
    console.log(pair[0] + ": " + pair[1]);
  }
});
   let txt = "";
for (let x in info) {
txt += info[x] + " ";
};

   document.write("a"");   
   document.write(info);   
   document.write(txt);   
   document.write("z");  
   
   document.write( window.location.pathname);
   alert("aaa"); 

</script> 
    <P />&nbsp; 
        <P />&nbsp; 
$={info}
  <%= info %> <P />&nbsp; 
        Bye
      <P />&nbsp; 
             
   </body>
</html>
