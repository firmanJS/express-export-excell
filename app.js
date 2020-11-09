var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
var fs = require('fs');

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');
var json2xls = require('json2xls');
var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.use(json2xls.middleware);
app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', indexRouter);
app.use('/users', usersRouter);
app.use('/excell',(req, res) => {
  const readJson = JSON.parse(fs.readFileSync(path.join("dummy.json"), "utf8"));
  res.set({
    'Content-Type': 'application/vnd.ms-excel',
    'Content-Disposition':'attachment;Filename=Excells.xls'
  })
  res.render('excell', {res:readJson});
})

app.use('/excell-via-package',(req, res) => {
  const readJson = JSON.parse(fs.readFileSync(path.join("dummy.json"), "utf8"));
  res.xls('data.xlsx', readJson.data);
})

// catch 404 and forward to error handler
app.use(function(req, res, next) {
  next(createError(404));
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
