<?php

header('Content-Type: text/html; charset=ISO-8859-1');

class Txt2xlsController extends BaseController {

    public function index()
    {
        return View::make('txt2xls/index');
    }

    public function convert()
    {
        $data = [];
        $file = Input::file('txt');
        $name = $this->clean($file->getClientOriginalName());

        $file->move('txt_temp/', $name);

        $file = fopen('txt_temp/' . $name, "r") or exit("No se pudo leer el archivo");

        while( ! feof( $file ) )
        {
            $line = trim( fgets($file) );
            array_push($data, $line);
        }

        fclose($file);

        foreach ($data as $key => $value) {
            $value = utf8_encode($value);
            $data[$key] = $this->clean($value);
        }

        Excel::create($name, function($excel) use ($name, $data) {

            $excel->setTitle($name);

            $excel->sheet('Observaciones acerca del layout');
            $excel->sheet('Encabezado', function($sheet) {
                $sheet->setColumnFormat([
                    'A:Q' => '@'
                ]);
            });
            $excel->sheet('Partidas', function($sheet) {
                $sheet->setColumnFormat([
                    'A' => '0.00',
                    'B' => '@',
                    'C' => '@',
                    'D' => '0.00'
                ]);
            });

            $excel->setActiveSheetIndex(1)
                    ->cell('A1', 'Cliente')
                    ->cell('A2', $data[1])
                    ->cell('B1', 'RFC')
                    ->cell('B2', $data[2])
                    ->cell('C1', 'Calle')
                    ->cell('C2', $data[3])
                    ->cell('D1', 'Num_Ext')
                    ->cell('E1', 'Num_Int')
                    ->cell('F1', 'Colonia')
                    ->cell('F2', $data[4])
                    ->cell('G1', 'Población')
                    ->cell('G2', $data[5])
                    ->cell('H1', 'Municipio')
                    ->cell('I1', 'Estado')
                    ->cell('J1', 'CP')
                    ->cell('J2', $data[6])
                    ->cell('K1', 'Cta_Predial')
                    ->cell('L1', 'IVA')
                    ->cell('M1', 'IVA_Ret')
                    ->cell('N1', 'ISR_Ret')
                    ->cell('O1', 'Impuesto_local')
                    ->cell('P1', 'IVA_Exento')
                    ->cell('Q1', 'Operacion_publico_general');

            $excel->setActiveSheetIndex(2)
                ->cell('A1', 'Cantidad')
                ->cell('B1', 'Producto')
                ->cell('C1', 'Unidad')
                ->cell('D1', 'Precio_Unitario');

            $sheet = $excel->setActiveSheetIndex(2);
            $row = 2;
            for ( $i = 9; $i < count($data); $i+=7 )
            {
                $line = $i;

                // cantidad                         9
                $sheet->cell("A$row", $data[$line]);

                // codigo de barras                 10
                $line++;
                $description = $data[$line] . ' ';

                // descripcion                      11
                $line++;
                $description .= $data[$line] . ' ';

                // observaciones                    12
                $line++;
                $description .= $data[$line] . ' ';

                // numeros de serie                 13
                $line++;
                if ( $data[$line] != '' ) {
                    $list = 'S/N ';
                    $series = explode('|', $data[$line]);
                    array_pop($series);
                    foreach ( $series as $sn ) {
                        $sn = trim($sn);
                        $list .= $sn . ', ';
                    }
                    $description .= $list;
                }
                $sheet->cell("B$row", $description);

                // unidad
                $sheet->cell("C$row", 'PIEZA');

                // costo unitario                   14
                $line++;
                $sheet->cell("D$row", $data[$line]);

                // linea en blanco                  15
                $line += 1;
                $row ++;
            }

            $excel->setActiveSheetIndex(0);

        })->download('xlsx');
    }

    public function clean($string)
    {

        $string = trim($string);

        $string = str_replace(
            array('á', 'à', 'ä', 'â', 'ª', 'Á', 'À', 'Â', 'Ä'),
            array('a', 'a', 'a', 'a', 'a', 'A', 'A', 'A', 'A'),
            $string
        );

        $string = str_replace(
            array('é', 'è', 'ë', 'ê', 'É', 'È', 'Ê', 'Ë'),
            array('e', 'e', 'e', 'e', 'E', 'E', 'E', 'E'),
            $string
        );

        $string = str_replace(
            array('í', 'ì', 'ï', 'î', 'Í', 'Ì', 'Ï', 'Î'),
            array('i', 'i', 'i', 'i', 'I', 'I', 'I', 'I'),
            $string
        );

        $string = str_replace(
            array('ó', 'ò', 'ö', 'ô', 'Ó', 'Ò', 'Ö', 'Ô'),
            array('o', 'o', 'o', 'o', 'O', 'O', 'O', 'O'),
            $string
        );

        $string = str_replace(
            array('ú', 'ù', 'ü', 'û', 'Ú', 'Ù', 'Û', 'Ü'),
            array('u', 'u', 'u', 'u', 'U', 'U', 'U', 'U'),
            $string
        );

        $string = str_replace(
            array('ñ', 'Ñ', 'ç', 'Ç'),
            array('n', 'N', 'c', 'C',),
            $string
        );

        //Esta parte se encarga de eliminar cualquier caracter extraño
        $string = str_replace(
            array("\\", "¨", "º", "-", "~",
                "#", "@", "|", "!", "\"",
                "·", "$", "%", "&", "/",
                "(", ")", "?", "'", "¡",
                "¿", "[", "^", "`", "]",
                "+", "}", "{", "¨", "´",
                ">", "< ", ";", ":",),
            '',
            $string
        );

        return $string;
    }

} 