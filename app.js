const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { MongoClient } = require('mongodb');

const app = express();
const port = 3000;

app.use(express.static('public'));

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const mongoURL = 'mongodb://localhost:27017';
const dbName = 'Form';
const collectionName = 'data';

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});


app.post("/upload", upload.single("excelFile"), async (req, res) => {
    try {
      const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
  
      const client = new MongoClient(mongoURL);
  
      try {
        await client.connect();
        const db = client.db(dbName);
        const collection = db.collection(collectionName);
        await collection.insertMany(jsonData);
  
        const data = await collection.find({}).toArray();
        res.json(data); // Send the JSON data as the response
        
      } finally {
        await client.close();
      }
    } catch (error) {
      res.status(500).send("Error uploading data: " + error.message);
    }
  });

  app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
  });
  