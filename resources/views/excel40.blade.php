@extends('layouts.app')

@section('content')
    <div id="main">
        <div class="container container-fruid">
            <div class="form-group page-header">
                <h2>
                    <label class="control-label">Mini Exceller</label>
                </h2>
            </div>
            <form method="post" enctype="multipart/form-data" action="{{url('upload40')}}" id="form">
                {{ csrf_field() }}
                <div class="form-row">
                    <div class="col-md-4 mb-3">
                        <label>下限:</label>
                        <input type="text" name="rlower" value="5">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>上限:</label>
                        <input type="text" name="rupper" value="30">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>R標準值:</label>
                        <input type="text" name="rstandard" value="1.82">
                    </div>
                </div>
                <div class="form-row">
                    <div class="col-md-4 mb-3">
                        <label>下限:</label>
                        <input type="text" name="glower" value="5">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>上限:</label>
                        <input type="text" name="gupper" value="12">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>G標準值:</label>
                        <input type="text" name="gstandard" value="2.32">
                    </div>
                </div>
                <div class="form-row">
                    <div class="col-md-4 mb-3">
                        <label>下限:</label>
                        <input type="text" name="blower" value="5">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>上限:</label>
                        <input type="text" name="bupper" value="12">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>B標準值:</label>
                        <input type="text" name="bstandard" value="2.5">
                    </div>
                </div>
                <div class="form-row">
                    <div class="col-md-4 mb-3">
                        <label>chip pitch:</label>
                        <input type="text" name="chip" value="0.2">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>x pixel pitch:</label>
                        <input type="text" name="xpixel" value="0.75">
                    </div>
                    <div class="col-md-4 mb-3">
                        <label>y pixel pitch:</label>
                        <input type="text" name="ypixel" value="0.75">
                    </div>
                </div>
                <div class="form-row">
                    <div class="col-md-12 mb-3">
                        <label>檔案上傳</label>
                        <input type="file" name="file" class="form-control-file">
                    </div>
                </div>
                <div class="form-row">
                    <div class="col-md-12 mb-3">
                        <button type="submit" class="btn btn-primary" id="submitbtn"> 提交</button>
                    </div>
                </div>
            </form>
        </div>


    </div>

@endsection

@section('script')
    <script>

        $('#submitbtn').click(function (e) {

            e.preventDefault(); // avoid to execute the actual submit of the form.

            let form = $('#form');
            let url = $('#form').attr('action');
            $.ajax( {
                url: url,
                type: 'POST',
                data: new FormData(form[0]),
                processData: false,
                contentType: false,
                beforeSend:function(){
                    $('#loadimg').show();
                },
                success:function(){
                    $('#loadimg').hide();
                    location.href="{{url('downloadpage40')}}"
                }
            } );
        } );



    </script>
@endsection
