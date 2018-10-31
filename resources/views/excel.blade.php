
<form method="post" enctype="multipart/form-data" action="{{url('upload')}}">
    {{ csrf_field() }}
    下限:-<input type="text" name="rlower">% 上限:<input name="rupper">% R標準值:<input type="text" name="rstandard"><br>
    下限:-<input type="text" name="glower">% 上限:<input name="gupper">% G標準值:<input type="text" name="gstandard"><br>
    下限:-<input type="text" name="blower">% 上限:<input name="bupper">% B標準值:<input type="text" name="bstandard"><br>
    chip pitch:<input type="text" name="chip"> pixel pitch:<input type="text" name="pixel">

    <input type="file" name="file">
    <button type="submit"> 提交 </button>
</form>
