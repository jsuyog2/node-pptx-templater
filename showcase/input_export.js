const path = require('path')
const { PPTXTemplater } = require('../src/index.js')
const { readFileSync, existsSync, createWriteStream, writeFileSync } = require('fs')
const { Readable, Writable } = require('stream')
const { finished } = require('stream/promises')

const SHOWCASE_DIR = __dirname
const TEMPLATES_DIR = path.join(SHOWCASE_DIR, 'templates')
const OUTPUT_DIR = path.join(SHOWCASE_DIR, 'output')

async function runDemos() {
    // // Load from PPTX file
    const pptx = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await loadPPTX(pptx, 'sample_pptx');

    // Load from ZIP file
    const pptxZip = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.zip'));
    await loadPPTX(pptxZip, 'sample_zip');

    // Load from Folder file
    const pptxFolder = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample'));
    await loadPPTX(pptxFolder, 'sample_folder');

    // Load from XML file
    const pptxXml = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.xml'));
    await loadPPTX(pptxXml, 'sample_xml');

    // Load from office PPTx file
    const pptxOffice = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'corporate_classifier_ppt.pptx'));
    await loadPPTX(pptxOffice, 'corporate_classifier_ppt');

    // Load from URL
    const pptxUrl = await PPTXTemplater.load(`https://github.com/jsuyog2/node-pptx-templater/raw/refs/heads/main/templates/sample.pptx`)
    await loadPPTX(pptxUrl, 'sample_url');

    // Load from Buffer
    const pptxBuffer = await PPTXTemplater.load(readFileSync(path.join(TEMPLATES_DIR, 'sample.pptx')))
    await loadPPTX(pptxBuffer, 'sample_buffer');

    // Extract PPTX file
    await extractPPT();

    // Build PPT
    await buildPPT();

    // create blank pptx
    await createPPT();

    // get pptx buffer
    const pptxBufferExport = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getPptxBuffer(pptxBufferExport, 'sample_buffer_pptx');

    // get pptx stream
    const pptxStream = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getPptxStream(pptxStream, 'sample_stream_pptx');

    // extractToZip
    const pptxExtract = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await extractToZip(pptxExtract, 'sample_extract_zip');

    // toResponse
    const pptxResponse = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToResponse(pptxResponse, 'sample_response_pptx');

    //extractToXml
    const pptxXmlExtract = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await extractToXml(pptxXmlExtract, 'sample_xml');

    //getToFastify
    const pptxFastify = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToFastify(pptxFastify, 'sample_fastify');

    //getToExpress
    const pptxExpress = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToExpress(pptxExpress, 'sample_express');

    // toBase64
    const pptxBase64 = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToBase64(pptxBase64, 'sample_base64');

    //toArrayBuffer
    const pptxArrayBuffer = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToArrayBuffer(pptxArrayBuffer, 'sample_array_buffer');

    //toFlatXml
    const pptxFlatXml = await PPTXTemplater.load(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await getToFlatXml(pptxFlatXml, 'sample_flat_xml');

    //preload template load form cache
    await PPTXTemplater.preload(path.join(TEMPLATES_DIR, 'sample.pptx'));
    await savePptxPreload('sample_preload_pptx');


    //clear cache
    await PPTXTemplater.clearCache();
    await savePptxPreload('sample_preload_pptx');


}



/*
  Loads PPTX file and replaces text
 */
async function loadPPTX(ppt, file) {
    console.log(`------loaded ${file}------`);

    ppt.useSlide(1)

    ppt.replaceText({
        '{{title}}': `${file}.pptx`
    })
    savePPTX(ppt, file);

    console.log(`File exported successfully: ${file}_output.pptx\n`);
}

/*
  Extracts PPTX file to a folder
 */
async function extractPPT() {
    console.log(`------Extracting PPTX file------`);
    if (existsSync(path.join(TEMPLATES_DIR, 'sample.pptx'))) {
        await PPTXTemplater.extractPptx(path.join(TEMPLATES_DIR, 'sample.pptx'), path.join(OUTPUT_DIR, 'sample_extracted'), { overwrite: true })
    }
    console.log(`File extracted successfully: sample_extracted\n`);
}

/*
  Builds PPTX file from a folder
 */
async function buildPPT() {
    console.log(`------Building PPTX file------`);
    if (existsSync(path.join(TEMPLATES_DIR, 'sample.pptx'))) {
        await PPTXTemplater.buildPptx(path.join(TEMPLATES_DIR, 'sample'), path.join(OUTPUT_DIR, 'sample_build.pptx'), { overwrite: true })
    }
    console.log(`File built successfully: sample_build.pptx\n`);
}

async function returnPptx(ppt, file) {
    console.log(`------loaded ${file}------`);

    ppt.useSlide(1)

    ppt.replaceText({
        '{{title}}': `${file}.pptx`
    })


    return ppt;
}

/* save pptx file */
async function savePPTX(ppt, file) {
    console.log(`------Saving PPTX file from ${file}------`);
    await ppt.saveToFile(path.join(OUTPUT_DIR, `${file}_output.pptx`));
    console.log(`File exported successfully: ${file}.pptx\n`);
}

/* get pptx buffer */
async function getPptxBuffer(pptx, file) {
    console.log(`------Getting PPTX buffer from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxBuffer = await ppt.toBuffer()
    await writeFileSync(path.join(OUTPUT_DIR, `${file}.pptx`), pptxBuffer);
    console.log(`📦 Buffer size: ${(pptxBuffer.length / 1024).toFixed(1)} KB`)
    console.log(`PPTX buffer successfully\n`);
    return pptxBuffer;
}

/* get pptx stream */
async function getPptxStream(pptx, file) {
    console.log(`------Getting PPTX stream from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxStream = await ppt.toStream()
    await pptxStream.pipe(createWriteStream(path.join(OUTPUT_DIR, `${file}.pptx`)));
    console.log(`File exported successfully: ${file}.pptx\n`);
    return pptxStream;
}

/* extractToZip
 */
async function extractToZip(pptx, file) {
    console.log(`------Extracting PPTX file from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    await ppt.extractToZip(path.join(OUTPUT_DIR, `${file}.zip`))
    console.log(`PPTX extracted successfully: ${file}.zip\n`);
}

/*extractToXml
*/
async function extractToXml(pptx, file) {
    console.log(`------Extracting PPTX file from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    await ppt.extractToXml(path.join(OUTPUT_DIR, `${file}`))
    console.log(`PPTX extracted successfully: ${file}\n`);
}

/* getToBase64 */
async function getToBase64(pptx, file) {
    console.log(`------Getting PPTX base64 from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxBase64 = await ppt.toBase64()
    const buffer = Buffer.from(pptxBase64, 'base64');
    await writeFileSync(path.join(OUTPUT_DIR, `${file}.pptx`), buffer);
    console.log(`Base64 successfully: ${file}.pptx\n`);
    return pptxBase64;
}

/* getToArrayBuffer 
*/
async function getToArrayBuffer(pptx, file) {
    console.log(`------Getting PPTX array buffer from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxArrayBuffer = await ppt.toArrayBuffer();

    const buffer = Buffer.from(pptxArrayBuffer);
    await writeFileSync(path.join(OUTPUT_DIR, `${file}.pptx`), buffer);
    console.log(`Array buffer successfully: ${file}.pptx\n`);
    return pptxArrayBuffer;
}

/* getToFlatXml
*/
async function getToFlatXml(pptx, file) {
    console.log(`------Getting PPTX flat xml from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxFlatXml = await ppt.toFlatXml()
    await writeFileSync(path.join(OUTPUT_DIR, `${file}.xml`), pptxFlatXml);
    console.log(`Flat xml successfully: ${file}.xml\n`);
    return pptxFlatXml;
}

/* save pptx preload and load form cache*/
async function savePptxPreload(file) {
    console.log(`------Saving PPTX file from ${file}------`);
    const ppt_preload = await PPTXTemplater.fromCache(path.join(TEMPLATES_DIR, 'sample.pptx'));
    const ppt = await returnPptx(ppt_preload, file)
    await ppt.saveToFile(path.join(OUTPUT_DIR, `${file}.pptx`));
    console.log(`File exported successfully: ${file}.pptx\n`);
}

/* 
toResponse*/
async function getToResponse(pptx, file) {
    console.log(`------Getting PPTX response from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    const pptxResponse = await ppt.toResponse()
    const destination = createWriteStream(path.join(OUTPUT_DIR, `${file}.pptx`));
    const nodeStream = Readable.fromWeb(pptxResponse.body);
    nodeStream.pipe(destination);
    await finished(destination);
    console.log(`PPTX response successfully: ${file}.pptx\n\n`);
}


/* Mock Response Object */
function mockRes() {
    const chunks = [];

    // 1. Assign the object to a variable first so we can attach event listeners
    const fakeRes = Object.assign(
        new Writable({
            write(chunk, encoding, callback) {
                chunks.push(chunk); // Capture data chunks sent by jszip
                callback();
            }
        }),
        {
            // --- Shared Core State ---
            statusCode: 200,
            headers: {},
            body: null,
            headersSent: false,

            // --- Routing / Status Chaining ---
            status(code) {
                this.statusCode = code;
                return this;
            },
            code(code) {
                this.statusCode = code;
                return this;
            },

            // --- Headers (Express + Fastify compatibility) ---
            set(key, value) {
                if (typeof key === 'object') {
                    Object.assign(this.headers, key);
                } else {
                    this.headers[key] = value;
                }
                return this;
            },
            header(key, value) {
                this.headers[key] = value;
                return this;
            },
            setHeader(key, value) {
                this.headers[key] = value;
                return this;
            },
            type(typeString) {
                this.headers['content-type'] = typeString;
                return this;
            },

            // --- Express/Fastify Response Delivery ---
            send(body) {
                if (body) {
                    if (Buffer.isBuffer(body) || typeof body === 'string') {
                        chunks.push(Buffer.from(body));
                    } else {
                        this.body = body; // Handle objects/JSON fallback
                    }
                }
                this.headersSent = true;
                this.end(); // Trigger the stream end
                return this;
            },
            json(body) {
                return this.send(body);
            }
        }
    );

    // 2. Attach the lifecycle listener BEFORE returning the object
    fakeRes.on('finish', () => {
        if (chunks.length > 0) {
            fakeRes.body = Buffer.concat(chunks);
        }
        // Asynchronously emit 'end' for stream consumers checking for completion
        process.nextTick(() => fakeRes.emit('end'));
    });

    // 3. Return the fully configured instance
    return fakeRes;
}


/* toFastify */
async function getToFastify(pptx, file) {
    console.log(`------Getting PPTX response from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    let res = mockRes();

    const streamFinished = new Promise((resolve, reject) => {
        res.on('finish', () => resolve());
        res.on('error', (err) => reject(err));
    });

    await ppt.toFastify(res, { filename: `${file}.pptx` })

    await streamFinished;
    await writeFileSync(path.join(OUTPUT_DIR, `${file}.pptx`), res.body);
    console.log(`PPTX fastify response successfully: ${file}.pptx\n\n`);
}

/* toExpress */
async function getToExpress(pptx, file) {
    console.log(`------Getting PPTX response from ${file}------`);
    const ppt = await returnPptx(pptx, file)
    let res = mockRes();

    const streamFinished = new Promise((resolve, reject) => {
        res.on('finish', () => resolve());
        res.on('error', (err) => reject(err));
    });

    await ppt.toExpress(res, { filename: `${file}.pptx` })

    await streamFinished;

    await writeFileSync(path.join(OUTPUT_DIR, `${file}.pptx`), res.body);
    console.log(`PPTX express response successfully: ${file}.pptx\n\n`);
}


/* 
    Create a pptx using blank pptx 
*/
async function createPPT() {
    console.log(`------Creating PPTX file------`);
    const ppt = await PPTXTemplater.createBlank();

    // Add a slide
    ppt.addSlide();

    ppt.useSlide(1);

    ppt.addShape('rect', {
        x: 100,
        y: 100,
        width: 100,
        height: 100,
        text: 'Hello World',
        fill: '#000000',
        textStyle: {
            color: '#ffffff'
        }
    })

    await ppt.saveToFile(path.join(OUTPUT_DIR, 'sample_blank_create.pptx'), { overwrite: true })
    console.log(`File created successfully: sample_blank_create.pptx\n`);
}


runDemos().catch(console.error);