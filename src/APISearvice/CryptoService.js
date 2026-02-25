import CryptoJS from "crypto-js";

const secretKey = "123456"; // Replace with your actual secret key

export const encryptId = (id) => {
  debugger
  return CryptoJS.AES.encrypt(id, secretKey).toString();
};

export const decryptId = (encryptedId) => {
  const bytes = CryptoJS.AES.decrypt(encryptedId, secretKey);
  return bytes.toString(CryptoJS.enc.Utf8);
};


