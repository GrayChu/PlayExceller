<!DOCTYPE html>
<html lang="{{ app()->getLocale() }}">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <!-- CSRF Token -->
    <meta name="csrf-token" content="{{ csrf_token() }}">

    <title>{{ config('app.name', 'Laravel') }}</title>

    <!-- Styles -->
    <link href="{{ asset('css/app.css') }}" rel="stylesheet">
    <script src="{{ asset('js/app.js') }}"></script>
    <script src="http://malsup.github.com/jquery.form.js"></script>

</head>
<body>
<div style="background-color:rgba(0,0,0,0.5);height: 100%;width: 100%;position: absolute;z-index: 1;display: none"
     id="loadimg"
     class="text-center">
    <div style="margin-top: 300px" class="row">
        <div class="cp-spinner cp-meter"></div><div class="text-white"><font color="white">loading...</font></div>
    </div>
</div>
<div id="app">
    <nav class="navbar  navbar-static-top navbar-default">
        <div class="container">
            <div class="navbar-header">
                <!-- Branding Image -->
                <a class="navbar-brand" href="{{ url('/') }}">
                    {{ config('app.name', 'Playnitride') }}
                </a>
            </div>
        </div>
    </nav>
    @yield('content')
</div>
<script>
    $.ajaxSetup({
        headers: {
            'X-CSRF-TOKEN': $('meta[name="csrf-token"]').attr('content')
        }
    });
</script>
@yield('script')
</body>

</html>
