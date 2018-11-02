@extends('layouts.app')

@section('content')
<div id="main">
  <div class="container container-fruid">
    <div class="form-group page-header">
       <h2>
           <label class="control-label">Exceller</label>
       </h2>
   </div>
    <form method="post" enctype="multipart/form-data" action="{{url('upload')}}">
      {{ csrf_field() }}
      <div class="form-row">
        <div class="col-md-4 mb-3">
          <label>下限:</label>
          <input type="text" name="rlower">
        </div>
        <div class="col-md-4 mb-3">
          <label>上限:</label>
          <input type="text" name="rupper">
        </div>
        <div class="col-md-4 mb-3">
          <label>R標準值:</label>
          <input type="text" name="rstandard">
        </div>
      </div>
      <div class="form-row">
        <div class="col-md-4 mb-3">
          <label>下限:</label>
          <input type="text" name="glower">
        </div>
        <div class="col-md-4 mb-3">
          <label>上限:</label>
          <input type="text" name="gupper">
        </div>
        <div class="col-md-4 mb-3">
          <label>G標準值:</label>
          <input type="text" name="gstandard">
        </div>
      </div>
      <div class="form-row">
        <div class="col-md-4 mb-3">
          <label>下限:</label>
          <input type="text" name="blower">
        </div>
        <div class="col-md-4 mb-3">
          <label>上限:</label>
          <input type="text" name="bupper">
        </div>
        <div class="col-md-4 mb-3">
          <label>B標準值:</label>
          <input type="text" name="bstandard">
        </div>
      </div>
      <div class="form-row">
        <div class="col-md-6 mb-3">
          <label>chip pitch:</label>
          <input type="text" name="chip">
        </div>
        <div class="col-md-6 mb-3">
          <label>pixel pitch:</label>
          <input type="text" name="pixel">
        </div>
      </div>
      <div class="form-row">
        <div class="col-md-12 mb-3">
          <label>檔案上傳</label>
          <input type="file" name="file"  class="form-control-file">
        </div>
      </div>
      <div class="form-row">
        <div class="col-md-12 mb-3">
          <button type="submit" class="btn btn-primary"> 提交 </button>
        </div>
      </div>
    </form>
  </div>
</div>
@endsection

@section('script')
@endsection
