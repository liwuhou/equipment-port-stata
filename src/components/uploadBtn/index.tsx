import { Button, Flex, message, Upload } from 'antd'
import { InboxOutlined } from '@ant-design/icons'
import type { UploadProps } from 'antd'

interface Prop {
  onUpload?: () => void
  onReset?: () => void
}

const { Dragger } = Upload

const DraggerUploadArea: React.FC<UploadProps> = (props) => {
  return (
    <Dragger {...props}>
      <p className="ant-upload-drag-icon">
        <InboxOutlined />
      </p>
      <p className="ant-upload-text">点击或拖入文件以上传</p>
    </Dragger>
  )
}

const UploadBtn: React.FC<Prop> = ({ onUpload, onReset }) => {
  const handleReset = () => {
    onReset?.()
  }

  const handleFileChange: UploadProps['onChange'] = (info) => {
    console.log('🤔 ~ info:', info)
    //
    onUpload?.()
  }

  return (
    <Flex gap="middle">
      <DraggerUploadArea onChange={handleFileChange} multiple={false} name="file" showUploadList={false} />
      <Button onClick={handleReset}>重置</Button>
    </Flex>
  )
}

export default UploadBtn
