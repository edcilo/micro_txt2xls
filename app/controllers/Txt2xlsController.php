<?php

class Txt2xlsController extends BaseController {

    public function index()
    {
        return View::make('txt2xls/index');
    }

    public function convert()
    {
        $data = [];
        $file = Input::file('txt');
        $name = $file->getClientOriginalName();

        $file->move('txt_temp/', $name);

        $file = fopen('txt_temp/' . $name, "r") or exit("No se pudo leer el archivo");

        while( ! feof( $file ) )
        {
            $line = trim( fgets($file) );
            array_push($data, $line);
        }

        fclose($file);

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
                ->setCellValue('A1', 'Cliente')
                ->setCellValue('A2', $data[1])
                ->setCellValue('B1', 'RFC')
                ->setCellValue('B2', $data[2])
                ->setCellValue('C1', 'Calle')
                ->setCellValue('C2', $data[3])
                ->setCellValue('D1', 'Num_Ext')
                ->setCellValue('E1', 'Num_Int')
                ->setCellValue('F1', 'Colonia')
                ->setCellValue('F2', $data[4])
                ->setCellValue('G1', 'Población')
                ->setCellValue('G2', $data[5])
                ->setCellValue('H1', 'Municipio')
                ->setCellValue('I1', 'Estado')
                ->setCellValue('J1', 'CP')
                ->setCellValue('J2', $data[6])
                ->setCellValue('K1', 'Cta_Predial')
                ->setCellValue('L1', 'IVA')
                ->setCellValue('M1', 'IVA_Ret')
                ->setCellValue('N1', 'ISR_Ret')
                ->setCellValue('O1', 'Impuesto_local')
                ->setCellValue('P1', 'IVA_Exento')
                ->setCellValue('Q1', 'Operacion_publico_general');

            $excel->setActiveSheetIndex(2)
                ->setCellvalue('A1', 'Cantidad')
                ->setCellvalue('B1', 'Producto')
                ->setCellvalue('C1', 'Unidad')
                ->setCellvalue('D1', 'Precio_Unitario');

            $row = 2;
            for ( $i = 9; $i < count($data); $i+=7 )
            {
                $line = $i;

                // cantidad                         9
                $excel->setCellvalue("A$row", $data[$line]);

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
                $excel->setCellvalue("B$row", $description);

                // unidad
                $excel->setCellvalue("C$row", 'PIEZA');

                // costo unitario                   14
                $line++;
                $excel->setCellvalue("D$row", $data[$line]);

                // linea en blanco                  15
                $line += 1;
                $row ++;
            }

            $excel->setActiveSheetIndex(0);

        })->download('xlsx');
    }

} 