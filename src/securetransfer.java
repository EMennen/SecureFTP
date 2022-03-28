import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.File;

import com.jcraft.jsch.*;
import com.jcraft.jsch.ChannelSftp.LsEntry;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;
import java.util.Date;
import java.io.*;
import java.net.*;
import java.lang.System.*;

/**
 * This program uses sFTP to transfer file's
 * defined in external textfile or all file's in a directory.
 * There is also the posiblity to remove file's after transfer.
 *
 * @author Erik Mennen
 *
 */
public class securetransfer {

    public static void main(String[] parameters) {
        String userName = parameters[0];
        String passWord = parameters[1];
        String hostName = parameters[2];
        String portS = parameters[3];
        int port = Integer.parseInt(portS);
        String fileName = parameters[4];
        String localPath = parameters[5];
        String remotePath = parameters[6];
        String action = parameters[7];
        String remove = parameters[8];

        String logmessage;
        Date dateStarted = new Date();
        long timeStarted = dateStarted.getTime();

        // if(!checkServer(hostName, port)){
        // logmessage = "Server" + hostName + ":" + port + "not found";
        // log2txt(logmessage);
        // System.exit(0);
        // }

        JSch jsch = new JSch();
        Session session = null;
        System.setProperty("user.dir", localPath);
        try {
            String fullName = "";
            String entryName = "";

            List<String> files = new ArrayList<String>();
            // Init sFTP session
            session = jsch.getSession(userName, hostName, port);
            session.setConfig("StrictHostKeyChecking", "no");
            session.setPassword(passWord);
            session.connect();
            Channel channel = session.openChannel("sftp");
            channel.connect();
            ChannelSftp sftpChannel = (ChannelSftp) channel;
            logmessage = String.valueOf(dateStarted);
            log2txt(logmessage);
            logmessage = "Connect :" + userName + "@" + hostName;
            log2txt(logmessage);
            // Change remote directory
            try {
                if (remotePath.equals(null) || remotePath.equals(" ") || remotePath.equals("")) {
                    // do nothing
                } else {
                    sftpChannel.cd(remotePath);
                }
            } catch (SftpException e) {
                logmessage = ("remotePath not valid " + remotePath);
                log2txt(logmessage);
            }

            // Get filename to process
            if (fileName.equals("*")) {
                if (action.equals("P")) {
                    files = getLocalEntries(localPath);
                }
                if (action.equals("G")) {
                    files = getRemoteEntries(sftpChannel, remotePath);
                }

            } else {
                files = getFileNames(fileName);
            }

            // Process files
            for (int i = 0; i < files.size(); i++) {

                entryName = files.get(i);
                if (action.equals("P")) {
                    fullName = localPath + entryName;
                    sftpChannel.put(fullName, entryName);
                    // System.out.println("Put " + fullName + " " + entryName);
                    logmessage = ("Put " + fullName + "  " + entryName);
                    log2txt(logmessage);
                    File localFile = new File(fullName);
                    if (remove.equals("Y")) {
                        if (localFile.delete()) {
                            logmessage = ("Deleted the file: " + localFile.getName());
                            log2txt(logmessage);
                        } else {
                            logmessage = ("Failed to delete the file." + localFile.getName());
                            log2txt(logmessage);
                        }
                    }

                } else if (action.equals("G")) {
                    fullName = localPath + entryName;
                    sftpChannel.get(entryName, fullName);
                    // System.out.println("Get " + entryName);
                    logmessage = ("Get " + entryName + "  " + fullName);
                    log2txt(logmessage);
                    if (remove.equals("Y")) {
                        sftpChannel.rm(entryName);
                        // System.out.println("RM " + entryName);
                        logmessage = ("RM " + entryName);
                        log2txt(logmessage);
                    }
                }

            }
            sftpChannel.exit();
            session.disconnect();
        } catch (JSchException e) {
            logmessage = ("JSchException :" + e);
            log2txt(logmessage);
        } catch (SftpException e) {
            logmessage = ("SftpException :" + e);
            log2txt(logmessage);
        } catch (IOException e) {
            logmessage = ("IOException :" + e);
            log2txt(logmessage);
        } catch (Exception e) {
            logmessage = ("Exception :" + e);
            log2txt(logmessage);
        }
        System.exit(0);
    }

    public static List<String> getLocalEntries(String inPath) {
        List<String> output = new ArrayList<String>();

        // Creates a new File instance by converting the given pathname string
        // into an abstract pathname
        File f = new File(inPath);
        String logmessage;
        // Populates the array with names of files and directories
        try {
            File[] files = f.listFiles();
            for (File file : files) {
                if (file.isDirectory() == false) {
                    output.add(file.getName().toString());
                    // System.out.println(file.getName().toString());
                }
            }
            if (output.size() == 0) {
                logmessage = ("getLocalEntries : nothing to transfer");
                log2txt(logmessage);
            }
        } catch (Exception e) {
            logmessage = ("getLocalEntries : " + e);
            log2txt(logmessage);
        }

        return output;
    }

    public static List<String> getRemoteEntries(ChannelSftp chan, String inPath) throws IOException {
        String logmessage;
        try {
            @SuppressWarnings("unchecked")
            List<LsEntry> result = new ArrayList<>((Vector<LsEntry>) chan.ls("*.*"));

            if (result != null) {
                List<String> output = new ArrayList<String>();
                for (LsEntry entry : result) {
                    if (".".equals(entry.getFilename()) || "..".equals(entry.getFilename())) {
                        continue;
                    }
                    output.add(entry.getFilename());
                    // System.out.println(entry.getFilename());
                    logmessage = ("getRemoteEntries : " + entry.getFilename());
                    log2txt(logmessage);
                }
                if (output.size() == 0) {
                    // logmessage = ("getRemoteEntries : Nothing to transfer");
                    // log2txt(logmessage);
                }
                return output;
            }
        } catch (SftpException e) {
            logmessage = ("getRemoteEntries : " + e);
            log2txt(logmessage);
        }
        return null;

    }

    public static List<String> getFileNames(String fileName) throws IOException {
        String line;
        String logmessage;
        try {
            List<String> files = new ArrayList<String>();
            FileReader reader = new FileReader(fileName);
            BufferedReader bufferedReader = new BufferedReader(reader);
            while ((line = bufferedReader.readLine()) != null) {
                files.add(line);
            }
            reader.close();
            return files;
        } catch (IOException e) {
            logmessage = ("getFileNames : " + e);
            log2txt(logmessage);

        }
        return null;
    }

    public static void log2txt(String logmessage) {

        List<String> aList = new ArrayList<String>();

        // for (int i = 0; i < inparms.length; i++) {
        // aList.add(inparms[i]);
        // }
        File file = new File("/tmp/Storage.txt");
        if (!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        try {
            FileWriter fw = new FileWriter(file, true);
            BufferedWriter bw = new BufferedWriter(fw);
            // for (int i = 0; i < aList.size(); i++) {
            // bw.write(aList.get(i).toString());
            // }
            bw.write(logmessage);
            bw.newLine();
            bw.flush();
            bw.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static boolean checkServer(String host, int port) {
        try (Socket s = new Socket(host, port)) {
            s.close();
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}