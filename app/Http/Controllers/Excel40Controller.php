<?php


namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Excel;
use Session;
use Response;
use Storage;

class Excel40Controller extends Controller
{
    public function getViewPage()
    {
        return view('excel40');
    }

    public function getDownloadPage()
    {
        return view('downloadpage40');
    }

    public function upload(Request $request)
    {
        $filepath = $request->file('file')->getRealPath();
        $filename = explode(".", $request->file('file')->getClientOriginalName())[0];

        $rupper = $request->input('rupper');
        $rlower = $request->input('rlower') * (-1);
        $rstd = $request->input('rstandard');
        $gupper = $request->input('gupper');
        $glower = $request->input('glower') * (-1);
        $gstd = $request->input('gstandard');
        $bupper = $request->input('bupper');
        $blower = $request->input('blower') * (-1);
        $bstd = $request->input('bstandard');
        $excel = Excel::load($filepath)->get();
        $odata = $excel->toArray();
        $title = $excel->getTitle();
        $result1 = array();
        $result2 = array();
        $rmvarr = array();//red measure value
        $gmvarr = array();//green measure value
        $bmvarr = array();//blue measure value
        $chip = $request->input('chip');
        $xpixel = $request->input('xpixel');
        $ypixel = $request->input('ypixel');
        foreach ($odata as $v) {
            if ($v['color'] == "R") {
                array_push($rmvarr, str_replace(" V", "", $v['measure_value']));
            } elseif ($v['color'] == "G") {
                array_push($gmvarr, str_replace(" V", "", $v['measure_value']));
            } elseif ($v['color'] == "B") {
                array_push($bmvarr, str_replace(" V", "", $v['measure_value']));
            }
        }
        if (!$rstd) {
            $rstd = $this->median($rmvarr);
        }
        if (!$gstd) {
            $gstd = $this->median($gmvarr);
        }
        if (!$bstd) {
            $bstd = $this->median($bmvarr);
        }
        $rmin = $rstd * (1 + $rlower / 100);
        $rmax = $rstd * (1 + $rupper / 100);
        $gmin = $gstd * (1 + $glower / 100);
        $gmax = $gstd * (1 + $gupper / 100);
        $bmin = $bstd * (1 + $blower / 100);
        $bmax = $bstd * (1 + $bupper / 100);

        array_push($result1, ['Group', 'Rows', 'Lines', 'Color', 'Standard Value', 'Hi Limit', 'Low Limit', 'Measure Value', 'Result']);
        foreach ($odata as $v) {
            if ($v['color'] == "R") {
                if (str_replace(" V", "", $v['measure_value']) <= $rmax && str_replace(" V", "", $v['measure_value']) >= $rmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $rstd,
                        "hi_limit" => $rmax,
                        "low_limit" => $rmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $rstd,
                        "hi_limit" => $rmax,
                        "low_limit" => $rmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            } elseif ($v['color'] == "G") {
                if (str_replace(" V", "", $v['measure_value']) <= $gmax && str_replace(" V", "", $v['measure_value']) >= $gmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $gstd,
                        "hi_limit" => $gmax,
                        "low_limit" => $gmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $gstd,
                        "hi_limit" => $gmax,
                        "low_limit" => $gmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            } elseif ($v['color'] == "B") {
                if (str_replace(" V", "", $v['measure_value']) <= $bmax && str_replace(" V", "", $v['measure_value']) >= $bmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $bstd,
                        "hi_limit" => $bmax,
                        "low_limit" => $bmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $bstd,
                        "hi_limit" => $bmax,
                        "low_limit" => $bmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            }
        }
        Session::put('excel1', $result1);
        Session::put('excel2', $result2);
        Session::put('chip', $chip);
        Session::put('xpixel', $xpixel);
        Session::put('ypixel', $ypixel);
        Session::put('fname', $filename);
        Session::put('title', $title);
        unset($excel);
        unset($odata);
        unset($title);
        return $this->getDownloadPage();
    }

    function arrSortObjsByKey($key, $order = 'DESC')
    {
        return function ($a, $b) use ($key, $order) {
            // Swap order if necessary
            if ($order == 'DESC') {
                list($a, $b) = array($b, $a);
            }
            // Check data type
            if (is_numeric($a[$key])) {
                return $a[$key] - $b[$key]; // compare numeric
            } else {
                return strnatcasecmp($a[$key], $b[$key]); // compare string
            }
        };
    }

    public function export1()
    {
        $result1 = Session::get('excel1');
        $fname = Session::get('fname');
        $title = Session::get('title');
        Excel::create($fname, function ($excel) use ($result1, $title) {
            $excel->sheet($title, function ($sheet) use ($result1) {
                $sheet->rows($result1);
                $sheet->freezeFirstRow();
            });
        })->export('xlsx');
    }

    public function export2()
    {
        $result2 = Session::get('excel2');
        $fname = Session::get('fname');
        $title = Session::get('title');
        $excel2 = Excel::create($fname . '-pdf', function ($excel) use ($result2, $title) {
            $excel->sheet($title, function ($sheet) use ($result2) {
                $rno = 0;
                $gno = 0;
                $bno = 0;

                $BStyle = array(
                    'borders' => array(
                        'outline' => array(
                            'style' => 'thin'
                        )
                    )
                );
                $numarr = array();
                $numarr2 = array();

                for ($i = 0; $i < 10; $i++) {
                    for ($j = 1; $j <= 16; $j++) {
                        array_push($numarr, "");
                        array_push($numarr, "");
                        array_push($numarr, $j);
                    }
                    if ($i != 9) {
                        array_push($numarr, "");
                    }
                }
                for ($i = 0; $i < 10; $i++) {
                    for ($j = 1; $j <= 16; $j++) {
                        array_push($numarr2, "");
                        array_push($numarr2, "");
                        array_push($numarr2, $i * 16 + $j);
                    }
                    if ($i != 9) {
                        array_push($numarr2, "");
                    }
                }
                $j = 0;
                for ($i = 0; $i < 160; $i++) {
                    if ($i % 16 == 0 && $i != 0) {
                        $j++;
                    }
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 1)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 1)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 1 + 1)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 1 + 1)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 32 + 1)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 32 + 1)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 63 + 1)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 63 + 1)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 94 + 1)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 94 + 1)->getCoordinate());

                    for ($z = 0; $z < 30; $z++) {
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 1 + 1 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 1 + 1 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 32 + 1 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 32 + 1 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 63 + 1 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 63 + 1 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 94 + 1 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 94 + 1 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                    }
//                    $sheet->setWidth($this->getNameFromNumber($i * 3 + 1 + 1 + $j+1),2);
//                    $sheet->setWidth($this->getNameFromNumber($i * 3 + 2 + 1 + $j+1),2);
//                    $sheet->setWidth($this->getNameFromNumber($i * 3 + 3 + 1 + $j+1),2);
                }
                $sheet->row(1, $numarr2);
                $sheet->row(1 + 1, $numarr);
                $sheet->row(32 + 1, $numarr);
                $sheet->row(63 + 1, $numarr);
                $sheet->row(94 + 1, $numarr);
                $z = 0;
                for ($j = 0; $j < 4; $j++) {
                    for ($i = 1; $i <= 30; $i++) {
                        $sheet->getCell($this->getNameFromNumber(1) . ($j * 30 + $i + 1 + 1 + $z))->setValue($j * 30 + $i);
                        $sheet->getCell($this->getNameFromNumber(2) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(51) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(100) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(149) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(198) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(247) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(296) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(345) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(394) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        $sheet->getCell($this->getNameFromNumber(443) . ($j * 30 + $i + 1 + 1 + $z))->setValue($i);
                        if ($i % 30 == 0) {
                            $z++;
                        }
                            $sheet->setHeight($j * 30 + $i + 1 + 1 + $z, 26);
                    }
                }

//                for ($i = 3; $i <= 1000; $i++) {
//                    if ($this->getNameFromNumber($i) == 'B' || $this->getNameFromNumber($i) == 'AY' || $this->getNameFromNumber($i) == 'CV'
//                        || $this->getNameFromNumber($i) == 'ES' || $this->getNameFromNumber($i) == 'GP'||$this->getNameFromNumber($i) == 'IM'
//                    ||$this->getNameFromNumber($i) == 'KJ'||$this->getNameFromNumber($i) == 'MG'||$this->getNameFromNumber($i) == 'OD'
//                        ||$this->getNameFromNumber($i) == 'QA') {
//                        $sheet->setWidth($this->getNameFromNumber($i), 7);
//                    } else {
//                        $sheet->setWidth($this->getNameFromNumber($i), 2);
//                    }
//                }

//
//
                if (!empty($result2)) {
                    foreach ($result2 as $v) {
                        if ($v['group'] % 10 == 0) {//group 10.20.30.40
                            if ($v['color'] == "R") {
                                $rno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(9 * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + 9 + 1, (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1)->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#FF0000");
                                });
                            } elseif ($v['color'] == "G") {
                                $gno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(9 * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + 9 + 1, (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1)->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#00FF11");
                                });
                            } elseif ($v['color'] == "B") {
                                $bno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(9 * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + 9 + 1, (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1)->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#002BFF");
                                });
                            }
                        } else { //group other
                            if ($v['color'] == "R") {
                                $rno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(
                                    (($v['group'] % 10) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + ($v['group'] % 10) - 1 + 1,
                                    (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1
                                )
                                    ->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#FF0000");
                                });
                            } elseif ($v['color'] == "G") {
                                $gno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(
                                    (($v['group'] % 10) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + ($v['group'] % 10) - 1 + 1,
                                    (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1
                                )
                                    ->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#00FF11");
                                });
                            } elseif ($v['color'] == "B") {
                                $bno++;
                                $sheet->cell($sheet->getCellByColumnAndRow(
                                    (($v['group'] % 10) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + ($v['group'] % 10) - 1 + 1,
                                    (ceil($v['group'] / 10) - 1) * 30 + $v['rows'] + ceil($v['group'] / 10) + 1
                                )
                                    ->getCoordinate(), function ($cell) use ($v) {
                                    $cell->setBackground("#002BFF");
                                });
                            }
                        }


                    }
                } else {
                    $sheet->rows($result2);
                }
                //printdate
                $sheet->mergeCells('SH1:SI20');
                $sheet->getCell('SH1')->setValue('Printed Date:');//printdate cell
                $sheet->cells('SH1:SI20', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '48',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                    $cells->setTextRotation(-90);
                });
                $sheet->setBorder('SH1', 'thin');
                //printdate input
                $sheet->mergeCells('SH21:SI40');
                $sheet->setBorder('SH21', 'thin');
                    //R
                    $sheet->mergeCells('RZ3:SA7');
                    $sheet->getCell('RZ3')->setValue('R');
                    $sheet->cells('RZ3:SA7', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '72',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                    });
                    $sheet->setBorder('RZ3', 'thin');
                    //R color
                    $sheet->mergeCells('RZ8:SA19');
                    $sheet->cells('RZ8:SA19', function ($cells) {
                        $cells->setBackground("#FF0000");
                    });
                    $sheet->setBorder('RZ8', 'thin');
                    //G
                    $sheet->mergeCells('SB3:SC7');
                    $sheet->getCell('SB3')->setValue('G');
                    $sheet->cells('SB3:SC7', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '72',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                    });
                    $sheet->setBorder('SB3', 'thin');
                    //G Color
                    $sheet->mergeCells('SB8:SC19');
                    $sheet->cells('SB8:SC19', function ($cells) {
                        $cells->setBackground("#00FF11");
                    });
                    $sheet->setBorder('SB8', 'thin');
                    //B
                    $sheet->mergeCells('SD3:SE7');
                    $sheet->getCell('SD3')->setValue('B');
                    $sheet->cells('SD3', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '72',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                    });
                    $sheet->setBorder('SD3', 'thin');
                    //B Color
                    $sheet->mergeCells('SD8:SE19');
                    $sheet->cells('SD8:SE19', function ($cells) {
                        $cells->setBackground("#002BFF");
                    });
                    $sheet->setBorder('SD8', 'thin');
                    //lot no
                    $sheet->mergeCells('SH43:SI62');
                    $sheet->getCell('SH43')->setValue('Lot No:');
                    $sheet->setBorder('SH43', 'thin');
                    $sheet->cells('SH43:SI62', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '48',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                        $cells->setTextRotation(-90);
                    });
                    //lot no input
                    $sheet->mergeCells('SH63:SI82');
                    $sheet->setBorder('SH63', 'thin');
//                    //L板
//                    $sheet->mergeCells('CE6:EM6');
//                    $sheet->getCell('CE6')->setValue('L板Bonding 需180度倒轉');
//                    $sheet->setBorder('CE6', 'thin');
//                    $sheet->cells('CE6:EM6', function ($cells) {
//                        $cells->setFont(
//                            array(
//                                'family' => 'Calibri',
//                                'size' => '72',
//                            )
//                        );
//                        $cells->setAlignment('center');
//                        $cells->setValignment('center');
//                    });
                    //name
                    $sheet->mergeCells('SH85:SI104');
                    $sheet->getCell('SH85')->setValue('Name:');
                    $sheet->setBorder('SH85', 'thin');
                    $sheet->cells('SH85:SI104', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '48',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                        $cells->setTextRotation(-90);
                    });
                    //name input
                    $sheet->mergeCells('SH105:SI124');
                    $sheet->setBorder('SH105', 'thin');
                    //date
                    $sheet->mergeCells('SF85:SG104');
                    $sheet->getCell('SF85')->setValue('Date:');
                    $sheet->setBorder('SF85', 'thin');
                    $sheet->cells('SF85:SG104', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '48',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                        $cells->setTextRotation(-90);
                    });
                    //date input
                    $sheet->mergeCells('SF105:SG124');
                    $sheet->setBorder('SF105', 'thin');
                    //defect record
                    $sheet->mergeCells('RZ33:SE39');
                    $sheet->getCell('RZ33')->setValue('Defect record');
                    $sheet->setBorder('RZ33', 'thin');
                    $sheet->cells('RZ33:SE39', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '48',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                    });
                    //R
                    $sheet->mergeCells('RZ40:SA44');
                    $sheet->getCell('RZ40')->setValue('R');
                    $sheet->setBorder('RZ40', 'thin');
                    //R input
                    $sheet->mergeCells('RZ45:SA56');
                    $sheet->getCell('RZ45')->setValue($rno);
                    $sheet->setBorder('RZ45', 'thin');
                    //G
                    $sheet->mergeCells('SB40:SC44');
                    $sheet->getCell('SB40')->setValue('G');
                    $sheet->setBorder('SB40', 'thin');
                    //G input
                    $sheet->mergeCells('SB45:SC56');
                    $sheet->getCell('SB45')->setValue($gno);
                    $sheet->setBorder('SB45', 'thin');
                    //B
                    $sheet->mergeCells('SD40:SE44');
                    $sheet->getCell('SD40')->setValue('B');
                    $sheet->setBorder('SD40', 'thin');
                    //B input
                    $sheet->mergeCells('SD45:SE56');
                    $sheet->getCell('SD45')->setValue($bno);
                    $sheet->setBorder('SD45', 'thin');
//
                    $sheet->cells('RZ40:SD45', function ($cells) {
                        $cells->setFont(
                            array(
                                'family' => 'Calibri',
                                'size' => '72',
                            )
                        );
                        $cells->setAlignment('center');
                        $cells->setValignment('center');
                    });
//
//                    $sheet->cells('A15:IL15', function ($cells) {
//                        $cells->setFont(
//                            array(
//                                'family' => 'Calibri',
//                                'size' => '28',
//                                'bold' => true
//                            )
//                        );
//                        $cells->setAlignment('center');
//                        $cells->setValignment('center');
//                    });
//                    $sheet->cells('A17:A98', function ($cells) {
//                        $cells->setFont(
//                            array(
//                                'family' => 'Calibri',
//                                'size' => '28',
//                                'bold' => true
//                            )
//                        );
//                        $cells->setAlignment('center');
//                        $cells->setValignment('center');
//                    });
//
//
//                    $sheet->cells('B16:IL98', function ($cells) {
//                        $cells->setFont(
//                            array(
//                                'family' => 'Calibri',
//                                'size' => '28',
//                            )
//                        );
//                        $cells->setAlignment('center');
//                        $cells->setValignment('center');
//                    });
//
                $sheet->getsheetView()->setzoomScale('60');
                $sheet->getsheetView()->setzoomScaleNormal('60');
                $sheet->setorientation('landscape');
//                $sheet->setfitToPage(true);
//                $sheet->setverticalCentered(true);
//                $sheet->setfitToWidth(true);
//                $sheet->setfitToHeight(true);
            });
        })->export('xlsx');
    }

    public function export3()
    {
        $result2 = Session::get('excel2');
        $chip = Session::get('chip');
        $xpixel = Session::get('xpixel');
        $ypixel = Session::get('ypixel');
        $fname = Session::get('fname');
        $content = "";
        $result3 = array();
        foreach ($result2 as $v) {
            if ($v['color'] == "R") {
                $l = $ypixel + $chip;
            } elseif ($v['color'] == "G") {
                $l = $ypixel;
            } elseif ($v['color'] == "B") {
                $l = $ypixel - $chip;
            }

            if ($v['group'] % 10 == 0) {//10.20.30.40
                $arr = ["color" => $v['color'], "x" => (-1) * $xpixel * 16 * 9 + $v['lines'] * (-1) * ($xpixel) + $xpixel,
                    "y" => (-1) * $ypixel * 30 * (ceil($v['group'] / 10) - 1) + $v['rows'] * (-1) * ($ypixel) + $l];
                array_push($result3, $arr);
            } else { //1.2.3.4.6.7.8.9
                $arr = ["color" => $v['color'], "x" => (-1) * ($xpixel) * 16 * ($v['group'] % 10 - 1) + (-1) * ($xpixel) * $v['lines'] + $xpixel,
                    "y" => (-1) * $ypixel * 30 * (ceil($v['group'] / 10) - 1) + $v['rows'] * (-1) * ($ypixel) + $l];
                array_push($result3, $arr);
            }

        }
        usort($result3, $this->arrSortObjsByKey('y', 'ASC'));
        foreach ($result3 as $v) {
            $content .= $v['color'] . "\t" . $v['x'] . "\t" . $v['y'];
            $content .= "\r\n";
        }
        Storage::put($fname, $content);
        return response()->download(storage_path('app/' . $fname), $fname . '.Lxy')->deleteFileAfterSend(true);;
    }

    public function median($numbers = array())
    {
        if (!is_array($numbers)) {
            $numbers = func_get_args();
        }
        rsort($numbers);
        $mid = (count($numbers) / 2);
        return ($mid % 2 != 0) ? $numbers{$mid - 1} : (($numbers{$mid - 1}) + $numbers{$mid}) / 2;
    }

    public function getNameFromNumber($num)
    {
        $numeric = ($num - 1) % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval(($num - 1) / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumber($num2) . $letter;
        } else {
            return $letter;
        }
    }

}
