// Apache POI oracle for encrypted OOXML. Decrypts one CFB-wrapped package with a password and, for Agile
// encryption, independently verifies the data-integrity HMAC over the whole EncryptedPackage stream. Reads
//   OfficeOracle <file> <password> [expectedPlaintextZip]
// and prints one JSON object: {ok, type, hashAlgorithm, startsPK, plaintextSize, hmacVerified, matchedPlaintext?, error?}.
// Built and run by poi.ts; kept here so the oracle lives with the tests. POI (Apache-2.0) jars are supplied on
// the classpath (FILEN_OFFICE_POI), never bundled.
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import javax.crypto.Mac;

import org.apache.poi.poifs.crypt.CryptoFunctions;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public final class OfficeOracle {
    public static void main(String[] args) {
        try {
            System.out.println(run(args[0], args[1], args.length > 2 ? args[2] : null));
        } catch (Throwable error) {
            System.out.println("{\"ok\":false,\"error\":" + json(error.getClass().getSimpleName() + ": " + error.getMessage()) + "}");
        }
    }

    private static String run(String file, String password, String expected) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(Paths.get(file).toFile(), true)) {
            EncryptionInfo info = new EncryptionInfo(fs);
            boolean agile = info.getEncryptionMode() == EncryptionMode.agile;
            Decryptor decryptor = Decryptor.getInstance(info);

            if (!decryptor.verifyPassword(password)) {
                return "{\"ok\":false,\"error\":\"wrongPassword\"}";
            }

            byte[] plaintext;
            try (InputStream stream = decryptor.getDataStream(fs)) {
                plaintext = readAll(stream);
            }

            boolean startsPK = plaintext.length >= 2 && plaintext[0] == 'P' && plaintext[1] == 'K';
            boolean hmacVerified = false;
            String hashName = "";

            if (agile) {
                HashAlgorithm hash = info.getHeader().getHashAlgorithm();
                hashName = hash.ecmaString;
                byte[] key = decryptor.getIntegrityHmacKey();
                byte[] value = decryptor.getIntegrityHmacValue();
                byte[] encryptedPackage = readStream(fs, "EncryptedPackage");
                Mac mac = CryptoFunctions.getMac(hash);
                mac.init(new javax.crypto.spec.SecretKeySpec(key, hash.jceHmacId));
                hmacVerified = Arrays.equals(mac.doFinal(encryptedPackage), value);
                if (!hmacVerified) {
                    return "{\"ok\":false,\"type\":\"agile\",\"error\":\"integrity\"}";
                }
            }

            String matched = "";
            if (expected != null) {
                boolean equal = Arrays.equals(plaintext, Files.readAllBytes(Paths.get(expected)));
                matched = ",\"matchedPlaintext\":" + equal;
            }

            return "{\"ok\":true,\"type\":" + json(agile ? "agile" : "standard") + ",\"hashAlgorithm\":" + json(hashName)
                    + ",\"startsPK\":" + startsPK + ",\"plaintextSize\":" + plaintext.length + ",\"hmacVerified\":" + hmacVerified
                    + matched + "}";
        }
    }

    private static byte[] readStream(POIFSFileSystem fs, String name) throws Exception {
        try (InputStream stream = fs.getRoot().createDocumentInputStream(name)) {
            return readAll(stream);
        }
    }

    private static byte[] readAll(InputStream stream) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        byte[] buffer = new byte[8192];
        int read;
        while ((read = stream.read(buffer)) != -1) out.write(buffer, 0, read);
        return out.toByteArray();
    }

    private static String json(String value) {
        if (value == null) return "null";
        StringBuilder out = new StringBuilder("\"");
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (c == '"' || c == '\\') out.append('\\').append(c);
            else if (c < 0x20) out.append(String.format("\\u%04x", (int) c));
            else out.append(c);
        }
        return out.append('"').toString();
    }
}
