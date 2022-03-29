<?php
$base_dir = dirname( __FILE__ ) . '/../Upload/' . $_GET["id"];
if(!is_dir($base_dir))
    mkdir($base_dir, 0777);
move_uploaded_file($_FILES["uploadfile"]["tmp_name"], $base_dir . '/' . $_FILES["uploadfile"]["name"]);
?>
