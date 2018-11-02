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
</head>
<body>
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
    @section('script')
    <!-- Scripts -->
    <script src="{{ asset('js/app.js') }}"></script>
    @show
</body>
</html>
