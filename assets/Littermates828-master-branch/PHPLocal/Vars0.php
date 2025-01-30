<!DOCTYPE html>
<html>
   <head>
      <title>PHP Test</title>
   </head>
   <body>
      <p>Hello World</p>
      <P />&nbsp; 
      <P />&nbsp; 
      <P />&nbsp; 
         Server 
      <P />&nbsp; 
      <P />&nbsp; 
         <?PHP
            foreach($_SERVER as $key_name => $key_value) {
                print $key_name." = ".$key_value."<br>";
            }
            
            ?>
      <P />&nbsp; 
      <P />&nbsp; 
      <P />&nbsp; 
         phpinfo()
      <P />&nbsp; 
      <P />&nbsp; 
         <?PHP phpinfo(); ?>
      <P />&nbsp; 
         Session 
      <P />&nbsp; 
      <P />&nbsp; 
         <?PHP
         $b = array(1, 1, 2, 3, 5, 8);

         $arr = get_defined_vars();
         
         // print $b
         print_r($arr["b"]);
         
         /* print path to the PHP interpreter (if used as a CGI)
          * e.g. /usr/local/bin/php */
         # echo $arr["_"];
         
         // print the command-line parameters if any
         # print_r($arr["argv"]);
         
         // print all the server vars
         # All ready done print_r($arr["_SERVER"]);
         
         // print all the available keys for the arrays of variables
         # print_r(array_keys(get_defined_vars()."<BR />"));
            print_r(array_keys(get_defined_vars()));

         /*
            foreach($_SESSION as $key_name => $key_value) {
                print $key_name." = ".$key_value."<br>";    
            }
           */
            ?>
   </body>
</html>
