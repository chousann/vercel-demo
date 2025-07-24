const express = require('express');
const multer = require('multer');
const cors = require('cors');
const helmet = require('helmet');
const compression = require('compression');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun } = require('docx');

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(cors());
app.use(helmet());
app.use(compression());
app.use(express.json());
app.use(express.static('uploads'));
app.use(express.static('downloads'));

// 配置文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = 'uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + '.pdf');
  }
});

const upload = multer({ 
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('只支持PDF文件'), false);
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB限制
  }
});

// 确保下载目录存在
const downloadDir = 'downloads';
if (!fs.existsSync(downloadDir)) {
  fs.mkdirSync(downloadDir, { recursive: true });
}

// 转换历史存储
let conversionHistory = [];

// API路由
app.post('/api/convert', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '没有上传文件' });
    }

    const pdfPath = req.file.path;
    const fileName = req.file.filename.replace('.pdf', '');
    
    console.log(`开始转换文件: ${fileName}`);

    // 读取PDF文件
    const pdfBuffer = fs.readFileSync(pdfPath);
    
    // 解析PDF内容
    const pdfData = await pdfParse(pdfBuffer);
    const text = pdfData.text;

    // 创建Word文档
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: text,
                size: 24,
                font: 'Arial'
              })
            ]
          })
        ]
      }]
    });

    // 生成Word文件
    const buffer = await Packer.toBuffer(doc);
    const wordFileName = `${fileName}.docx`;
    const wordPath = path.join(downloadDir, wordFileName);
    
    fs.writeFileSync(wordPath, buffer);

    // 记录转换历史
    const historyItem = {
      id: Date.now().toString(),
      originalName: req.file.originalname,
      fileName: wordFileName,
      status: 'completed',
      downloadUrl: `/downloads/${wordFileName}`,
      createdAt: new Date()
    };
    
    conversionHistory.unshift(historyItem);

    // 清理上传的PDF文件
    fs.unlinkSync(pdfPath);

    console.log(`转换完成: ${wordFileName}`);
    
    res.json({
      success: true,
      message: '转换成功',
      downloadUrl: `/downloads/${wordFileName}`,
      fileName: wordFileName
    });

  } catch (error) {
    console.error('转换失败:', error);
    
    // 记录失败历史
    const historyItem = {
      id: Date.now().toString(),
      originalName: req.file?.originalname || 'unknown',
      status: 'failed',
      error: error.message,
      createdAt: new Date()
    };
    
    conversionHistory.unshift(historyItem);

    res.status(500).json({
      success: false,
      error: '转换失败',
      message: error.message
    });
  }
});

// 获取转换历史
app.get('/api/history', (req, res) => {
  res.json(conversionHistory.slice(0, 20)); // 只返回最近20条记录
});

// 下载文件
app.get('/downloads/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(downloadDir, filename);
  
  if (fs.existsSync(filePath)) {
    res.download(filePath);
  } else {
    res.status(404).json({ error: '文件不存在' });
  }
});

// 健康检查
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// 错误处理中间件
app.use((error, req, res, next) => {
  console.error('服务器错误:', error);
  res.status(500).json({
    error: '服务器内部错误',
    message: error.message
  });
});

if (process.env.VERCEL) {
  module.exports = app;
} else {
  app.listen(PORT, () => {
    console.log(`PDF转Word服务器运行在端口 ${PORT}`);
    console.log(`API地址: http://localhost:${PORT}/api`);
  });
} 