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

Route::get('excel','ExcelController@getViewPage');
Route::get('excel40','Excel40Controller@getViewPage');
Route::post('/upload','ExcelController@upload');
Route::post('/upload40','Excel40Controller@upload');

Route::get('/downloadpage','ExcelController@getDownloadPage');
Route::get('export1','ExcelController@export1');
Route::get('export2','ExcelController@export2');
Route::get('export3','ExcelController@export3');


Route::get('/downloadpage40','Excel40Controller@getDownloadPage');
Route::get('/40/export1','Excel40Controller@export1');
Route::get('/40/export2','Excel40Controller@export2');
Route::get('/40/export3','Excel40Controller@export3');
Route::get('/40/export4R','Excel40Controller@export4R');
Route::get('/40/export4G','Excel40Controller@export4G');
Route::get('/40/export4B','Excel40Controller@export4B');

Route::get('/test',function(){


});