import os
import stat
import paramiko
import traceback


class SSH(object):

    def __init__(self, ip, port=22, username=None, password=None, timeout=30):
        self.ip = ip  # ssh远程连接的服务器ip
        self.port = 22  # ssh的端口一般默认是22，
        self.username = username  # 服务器用户名
        self.password = password  # 密码
        self.timeout = timeout  # 连接超时

        # paramiko.SSHClient() 创建一个ssh对象，用于ssh登录以及执行操作
        self.ssh = paramiko.SSHClient()

        # paramiko.Transport()创建一个文件传输对象，用于实现文件的传输
        self.t = paramiko.Transport(sock=(self.ip, self.port))

    def connect(self):
        try:
            self._password_connect()  # 密码登录
        except:
            print('ssh password connect faild!')

    def _password_connect(self):

        self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        self.ssh.connect(hostname=self.ip, port=22, username=self.username, password=self.password)
        self.t.connect(username=self.username, password=self.password)

    def close(self):
        self.t.close()  # 断开文件传输的连接
        self.ssh.close()  # 断开ssh连接

    def execute_cmd(self, cmd):

        stdin, stdout, stderr = self.ssh.exec_command(cmd)

        res, err = stdout.read(), stderr.read()
        result = res if res else err

        return result.decode()

    # 从远程服务器获取文件到本地
    def sftp_get(self, remotefile, localfile):

        sftp = paramiko.SFTPClient.from_transport(self.t)
        sftp.get(remotefile, localfile)

    # 从本地上传文件到远程服务器
    def sftp_put(self, localfile, remotefile):

        sftp = paramiko.SFTPClient.from_transport(self.t)
        sftp.put(localfile, remotefile)

    # 递归遍历远程服务器指定目录下的所有文件
    def _get_all_files_in_remote_dir(self, sftp, remote_dir):
        all_files = list()
        if remote_dir[-1] == '/':
            remote_dir = remote_dir[0:-1]

        files = sftp.listdir_attr(remote_dir)
        for file in files:
            filename = remote_dir + '/' + file.filename

            if stat.S_ISDIR(file.st_mode):  # 如果是文件夹的话递归处理
                all_files.extend(self._get_all_files_in_remote_dir(sftp, filename))
            else:
                all_files.append(filename)

        return all_files

    # 本地文件夹的上传到远程服务器
    def sftp_get_dir(self, remote_dir, local_dir):
        try:

            sftp = paramiko.SFTPClient.from_transport(self.t)

            all_files = self._get_all_files_in_remote_dir(sftp, remote_dir)

            for file in all_files:

                local_filename = file.replace(remote_dir, local_dir)
                local_filepath = os.path.dirname(local_filename)

                if not os.path.exists(local_filepath):
                    os.makedirs(local_filepath)

                sftp.get(file, local_filename)

        except:
            print('ssh get dir from master failed.')
            print(traceback.format_exc())  # 具体报错信息

    # 递归遍历本地服务器指定目录下的所有文件
    def _get_all_files_in_local_dir(self, local_dir):
        all_files = list()

        for root, dirs, files in os.walk(local_dir, topdown=True):
            for file in files:
                filename = os.path.join(root, file)
                all_files.append(filename)

        return all_files

    def sftp_put_dir(self, local_dir, remote_dir):
        try:
            sftp = paramiko.SFTPClient.from_transport(self.t)

            if remote_dir[-1] == "/":
                remote_dir = remote_dir[0:-1]

            all_files = self._get_all_files_in_local_dir(local_dir)
            for file in all_files:

                remote_filename = file.replace(local_dir, remote_dir)
                remote_path = os.path.dirname(remote_filename)

                try:
                    sftp.stat(remote_path)
                except:
                    # os.popen('mkdir -p %s' % remote_path)  这个bug，是在本地创建文件夹，晕了
                    self.execute_cmd('mkdir -p %s' % remote_path) # 使用这个远程执行命令

                sftp.put(file, remote_filename)

        except:
            print('ssh get dir from master failed.')
            print(traceback.format_exc())