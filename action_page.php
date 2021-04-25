<?php
    $command = escapeshellcmd("C:\Users\emuramoto1\Documents\GitHub\pssd\pssd\backend.py");
    $output = shell_exec($command);
    echo $output;
?>