@extends('layouts.app')

@section('content')
<div class="container container-fruid">
  <div class="form-group page-header">
     <h2>
         <label class="control-label">Download</label>
     </h2>
  </div>
  <div class="row">
    <div class="col-md-4 mb-3">
      <label>Requirement1</label>
      <button onclick="window.location.href='{{url('/export1')}}'" class="btn brn-primary">下載</button>
    </div>
    <div class="col-md-4 mb-3">
      <label>Requirement2</label>
      <button onclick="window.location.href='{{url('/export2')}}'" class="btn brn-primary">下載</button>
    </div>
    <div class="col-md-4 mb-3">
      <label>Requirement3</label>
      <button onclick="window.location.href='{{url('/export3')}}'" class="btn brn-primary">下載</button>
    </div>
  </div>
  <div class="row">
    <div class="col-md-4 mb-3">
      <label>Requirement4-R</label>
      <button onclick="window.location.href='{{url('/export4R')}}'" class="btn brn-primary">下載</button>
    </div>
    <div class="col-md-4 mb-3">
      <label>Requirement4-G</label>
      <button onclick="window.location.href='{{url('/export4G')}}'" class="btn brn-primary">下載</button>
    </div>
    <div class="col-md-4 mb-3">
      <label>Requirement4-B</label>
      <button onclick="window.location.href='{{url('/export4B')}}'" class="btn brn-primary">下載</button>
    </div>
  </div>
</div>
@endsection

@section('script')
@endsection
