/** Opens password-protected PDFs: decrypts in the engine worker, asking for a password when needed. */
import { isEncryptedPdf } from '../core/pdf-decrypt';
import { engine, type EngineError } from './engine/client';
import { promptText } from './ui';

/** Returns the PDF without encryption (the same bytes when it was not encrypted). Throws if the user cancels. */
export async function unlockPdf(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  if (!isEncryptedPdf(bytes)) return bytes;
  let password = '';
  for (let attempt = 0; attempt < 6; attempt++) {
    try {
      const copy = bytes.slice();
      return await engine.run<Uint8Array>({ op: 'decryptPdf', bytes: copy, password }, { transfer: [copy.buffer] });
    } catch (e) {
      const code = (e as EngineError).code;
      if (code !== 'password_required' && code !== 'password_invalid') throw e;
      const pw = await promptText(code === 'password_invalid' ? 'Неверный пароль' : 'PDF защищён паролем', `Пароль для «${name}»`, '', '', 'password');
      if (pw === null) throw new Error(`${name}: PDF защищён паролем — файл не открыт.`);
      password = pw;
    }
  }
  throw new Error(`${name}: слишком много неверных попыток ввода пароля.`);
}
