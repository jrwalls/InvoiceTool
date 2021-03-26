const express = require('express');
const app = express();
const cors = require('cors');
var exphbs = require('express-handlebars');
app.use(express.static('public'));

app.use(cors())
app.engine('handlebars', exphbs());
app.set('view engine', 'handlebars');

app.use(express.json({ limit: '50mb' }));


app.get('/', (req, res) => {
    res.render('home');
});

app.listen(5000);


