<?php
// Exporta los datasets PHP del tema MCR a JSON para que build_master.py los lea.
// Uso: php export_theme_data.php <ruta_al_tema> <salida.json>
define('ABSPATH', '/');
[$_, $theme, $out] = $argv;
$data = [
  'private'  => include "$theme/data/creators-private.php",
  'public'   => include "$theme/data/creators-public.php",
];
file_put_contents($out, json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
echo "ok\n";
