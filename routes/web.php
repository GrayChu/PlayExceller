<?php

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});

Route::get('/excel','ExcelController@getViewPage');
Route::post('/upload','ExcelController@upload');

Route::get('/downloadpage','ExcelController@getDownloadPage');
Route::get('export1','ExcelController@export1');
Route::get('export2','ExcelController@export2');
Route::get('export3','ExcelController@export3');
