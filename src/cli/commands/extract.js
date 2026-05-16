/**
 * @fileoverview `extract` CLI command — extracts XML parts from a PPTX.
 */
import chalk from 'chalk';
import { resolve } from 'path';
import { writeFileSync } from 'fs';
import { PPTXTemplater } from '../../index.js';

export async function extractCommand(filePath, opts) {
  try {
    const ppt = await PPTXTemplater.load(resolve(filePath));

    if (opts.slide) {
      const slideNum = parseInt(opts.slide, 10);
      // Access internal zip via the engine's buffer
      const buffer = await ppt.toBuffer();
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(buffer);
      const slideFile = zip.file(`ppt/slides/slide${slideNum}.xml`);

      if (!slideFile) {
        console.error(chalk.red(`Slide ${slideNum} not found`));
        process.exit(1);
      }

      const xml = await slideFile.async('text');

      if (opts.out) {
        writeFileSync(resolve(opts.out), xml, 'utf-8');
        console.log(chalk.green(`✓ Extracted slide ${slideNum} to ${opts.out}`));
      } else {
        console.log(xml);
      }
    } else {
      console.log(chalk.yellow('Specify --slide <number> to extract'));
      process.exit(1);
    }
  } catch (err) {
    console.error(chalk.red(`Extract failed: ${err.message}`));
    process.exit(1);
  }
}
