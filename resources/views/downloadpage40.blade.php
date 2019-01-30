@extends('layouts.app')

@section('content')
<div class="container container-fruid">
  <div class="form-group page-header">
     <h2>
         <label class="control-label">Download</label>
     </h2>
  </div>
  <div class="col-md-4 mb-3">
    <label>Requirement1</label>
    <button onclick="window.location.href='{{url('/40/export1')}}'" class="btn brn-primary">下載</button>
  </div>
  <div class="col-md-4 mb-3">
    <label>Requirement2</label>
    <button onclick="window.location.href='{{url('/40/export2')}}'" class="btn brn-primary">下載</button>
  </div>
  <div class="col-md-4 mb-3">
    <label>Requirement3</label>
    <button onclick="window.location.href='{{url('/40/export3')}}'" class="btn brn-primary">下載</button>
  </div>
</div>
@endsection

@section('script')
@endsection
