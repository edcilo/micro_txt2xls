

{{ Form::open(['route'=>'convert', 'files'=>'true']) }}

{{ Form::label('txt', 'Seleccionar archivo txt') }}
{{ Form::file('txt') }}

{{ Form::submit('Subir') }}

{{ Form::close() }}