import {
  Alert,
  Button,
  Card,
  Col,
  DatePicker,
  Form,
  Input,
  Row,
  Select,
  Space,
  Table,
  Tag,
  Typography,
  Upload,
  message,
} from 'antd';
import type { ColumnsType } from 'antd/es/table';
import type { UploadFile } from 'antd/es/upload/interface';
import {
  AudioOutlined,
  CheckCircleOutlined,
  FileTextOutlined,
  MobileOutlined,
  SendOutlined,
  UploadOutlined,
} from '@ant-design/icons';
import axios from 'axios';
import dayjs from 'dayjs';
import { useState } from 'react';

import {
  createMinutesFromUpload,
  createVisitVoc,
  type MeetingMinutes,
  type VisitVoc,
  type VisitVocPayload,
} from '../api/voc';

type VisitVocForm = Omit<VisitVocPayload, 'visit_date'> & {
  visit_date: dayjs.Dayjs;
};

type MinutesForm = {
  title: string;
  company_name?: string;
  meeting_date?: dayjs.Dayjs;
  participants?: string;
  transcript_text?: string;
};

const vocColumns: ColumnsType<VisitVoc> = [
  { title: '방문일', dataIndex: 'visit_date', key: 'visit_date', width: 110 },
  { title: '고객사', dataIndex: 'company_name', key: 'company_name' },
  { title: '담당자', dataIndex: 'visitor_name', key: 'visitor_name', width: 110 },
  { title: '분야', dataIndex: 'product_area', key: 'product_area', width: 120, responsive: ['md'] },
  {
    title: '감정',
    dataIndex: 'sentiment',
    key: 'sentiment',
    width: 90,
    render: (value) => <Tag color={value === '불만' ? 'error' : value === '긍정' ? 'success' : 'processing'}>{value}</Tag>,
  },
  { title: 'VOC', dataIndex: 'voc_text', key: 'voc_text', ellipsis: true },
  { title: '후속조치', dataIndex: 'next_action', key: 'next_action', ellipsis: true, responsive: ['lg'] },
];

function MinutesList({ title, items }: { title: string; items: string[] }) {
  return (
    <div className="minutes-block">
      <Typography.Text strong>{title}</Typography.Text>
      <ul>
        {items.map((item) => (
          <li key={item}>{item}</li>
        ))}
      </ul>
    </div>
  );
}

export function VisitVocPage() {
  const [vocForm] = Form.useForm<VisitVocForm>();
  const [minutesForm] = Form.useForm<MinutesForm>();
  const [recentVoc, setRecentVoc] = useState<VisitVoc[]>([]);
  const [minutes, setMinutes] = useState<MeetingMinutes | null>(null);
  const [recordingFiles, setRecordingFiles] = useState<UploadFile[]>([]);
  const [savingVoc, setSavingVoc] = useState(false);
  const [creatingMinutes, setCreatingMinutes] = useState(false);

  const handleCreateVoc = async (values: VisitVocForm) => {
    setSavingVoc(true);
    try {
      const payload: VisitVocPayload = {
        ...values,
        visit_date: values.visit_date.format('YYYY-MM-DD'),
      };
      const created = await createVisitVoc(payload);
      setRecentVoc((prev) => [created, ...prev].slice(0, 10));
      vocForm.resetFields();
      vocForm.setFieldValue('visit_date', dayjs());
      message.success('방문 VOC가 수집되었습니다.');
    } catch {
      message.error('VOC 입력 내용을 저장하지 못했습니다.');
    } finally {
      setSavingVoc(false);
    }
  };

  const handleCreateMinutes = async (values: MinutesForm) => {
    setCreatingMinutes(true);
    try {
      const file = recordingFiles[0]?.originFileObj;
      const data = await createMinutesFromUpload({
        ...values,
        meeting_date: values.meeting_date?.format('YYYY-MM-DD'),
        recording_file: file,
      });
      setMinutes(data);
      message.success('회의록 초안이 생성되었습니다.');
    } catch (error: unknown) {
      const detail = axios.isAxiosError(error) ? error.response?.data?.detail : undefined;
      message.error(typeof detail === 'string' ? detail : '회의록을 생성하지 못했습니다.');
    } finally {
      setCreatingMinutes(false);
    }
  };

  return (
    <Space className="dashboard-page visit-voc-page" direction="vertical" size={18}>
      <div className="page-heading">
        <div>
          <Typography.Title level={3}>방문 VOC 수집</Typography.Title>
          <Typography.Text type="secondary">현장 방문조직이 모바일에서 바로 입력하고 녹취 텍스트를 회의록으로 정리합니다.</Typography.Text>
        </div>
        <Space wrap>
          <Tag icon={<MobileOutlined />} color="blue">모바일 1분 입력</Tag>
          <Tag icon={<AudioOutlined />} color="purple">녹취 텍스트 회의록</Tag>
        </Space>
      </div>

      <Row gutter={[16, 16]}>
        <Col xs={24} lg={9}>
          <Card
            title={
              <Space>
                <MobileOutlined />
                1분 간편 VOC 입력
              </Space>
            }
          >
            <Form
              form={vocForm}
              layout="vertical"
              initialValues={{ visit_date: dayjs(), channel: '방문', sentiment: '보통' }}
              onFinish={handleCreateVoc}
            >
              <Form.Item name="company_name" label="고객사" rules={[{ required: true }]}>
                <Input placeholder="예: 하나투어" />
              </Form.Item>
              <Row gutter={12}>
                <Col xs={24} md={12}>
                  <Form.Item name="visit_date" label="방문일" rules={[{ required: true }]}>
                    <DatePicker className="full-width" />
                  </Form.Item>
                </Col>
                <Col xs={24} md={12}>
                  <Form.Item name="visitor_name" label="방문자" rules={[{ required: true }]}>
                    <Input placeholder="담당자명" />
                  </Form.Item>
                </Col>
              </Row>
              <Row gutter={12}>
                <Col xs={24} md={12}>
                  <Form.Item name="contact_name" label="고객 담당자">
                    <Input placeholder="면담자명" />
                  </Form.Item>
                </Col>
                <Col xs={24} md={12}>
                  <Form.Item name="channel" label="채널">
                    <Select
                      options={[
                        { label: '방문', value: '방문' },
                        { label: '통화', value: '통화' },
                        { label: '화상회의', value: '화상회의' },
                      ]}
                    />
                  </Form.Item>
                </Col>
              </Row>
              <Row gutter={12}>
                <Col xs={24} md={12}>
                  <Form.Item name="product_area" label="업무/상품">
                    <Input placeholder="CMS, ERP 연계 등" />
                  </Form.Item>
                </Col>
                <Col xs={24} md={12}>
                  <Form.Item name="sentiment" label="고객 반응">
                    <Select
                      options={[
                        { label: '보통', value: '보통' },
                        { label: '긍정', value: '긍정' },
                        { label: '불만', value: '불만' },
                      ]}
                    />
                  </Form.Item>
                </Col>
              </Row>
              <Form.Item name="voc_text" label="VOC 내용" rules={[{ required: true }]}>
                <Input.TextArea rows={4} placeholder="고객 요청, 불편, 개선 의견을 짧게 입력" />
              </Form.Item>
              <Form.Item name="next_action" label="후속조치">
                <Input.TextArea rows={2} placeholder="담당자 배정, 회신 기한, 내부 공유 내용" />
              </Form.Item>
              <Button type="primary" htmlType="submit" icon={<SendOutlined />} loading={savingVoc} block>
                VOC 등록
              </Button>
            </Form>
          </Card>
        </Col>

        <Col xs={24} lg={15}>
          <Card
            title={
              <Space>
                <FileTextOutlined />
                녹취 텍스트 회의록 변환
              </Space>
            }
          >
            <Alert
              className="compact-alert"
              showIcon
              type="info"
              message="통화녹음/화상회의 파일은 사내 STT 또는 Whisper 연동 후 텍스트를 함께 넣으면 회의록으로 정리됩니다."
            />
            <Form form={minutesForm} layout="vertical" onFinish={handleCreateMinutes}>
              <Row gutter={12}>
                <Col xs={24} md={12}>
                  <Form.Item name="title" label="회의명" rules={[{ required: true }]}>
                    <Input placeholder="고객 방문 미팅" />
                  </Form.Item>
                </Col>
                <Col xs={24} md={12}>
                  <Form.Item name="company_name" label="고객사">
                    <Input placeholder="고객사명" />
                  </Form.Item>
                </Col>
              </Row>
              <Row gutter={12}>
                <Col xs={24} md={12}>
                  <Form.Item name="meeting_date" label="회의일">
                    <DatePicker className="full-width" />
                  </Form.Item>
                </Col>
                <Col xs={24} md={12}>
                  <Form.Item name="participants" label="참석자">
                    <Input placeholder="홍길동, 김하나" />
                  </Form.Item>
                </Col>
              </Row>
              <Form.Item label="녹취/자막 파일">
                <Upload
                  accept=".txt,.md,.srt,.vtt,.csv,.mp3,.m4a,.wav,.mp4,.webm"
                  beforeUpload={() => false}
                  fileList={recordingFiles}
                  maxCount={1}
                  onChange={({ fileList }) => setRecordingFiles(fileList.slice(-1))}
                >
                  <Button icon={<UploadOutlined />}>파일 선택</Button>
                </Upload>
              </Form.Item>
              <Form.Item name="transcript_text" label="변환 텍스트">
                <Input.TextArea rows={7} placeholder="STT 변환 결과 또는 회의 자막 텍스트를 붙여넣기" />
              </Form.Item>
              <Button type="primary" htmlType="submit" icon={<CheckCircleOutlined />} loading={creatingMinutes}>
                회의록 생성
              </Button>
            </Form>
          </Card>
        </Col>
      </Row>

      {minutes ? (
        <Card title={minutes.title} extra={minutes.source_file_name ? `원본: ${minutes.source_file_name}` : null}>
          <Space direction="vertical" size={14} className="full-width">
            <Typography.Paragraph className="minutes-summary">{minutes.summary}</Typography.Paragraph>
            <Row gutter={[16, 16]}>
              <Col xs={24} md={12}>
                <MinutesList title="주요 VOC/안건" items={minutes.key_topics} />
              </Col>
              <Col xs={24} md={12}>
                <MinutesList title="결정사항" items={minutes.decisions} />
              </Col>
              <Col xs={24} md={12}>
                <MinutesList title="후속조치" items={minutes.action_items} />
              </Col>
              <Col xs={24} md={12}>
                <MinutesList title="리스크" items={minutes.risks} />
              </Col>
            </Row>
          </Space>
        </Card>
      ) : null}

      <Card title="최근 입력 VOC">
        <Table
          rowKey="id"
          columns={vocColumns}
          dataSource={recentVoc}
          pagination={{ pageSize: 5 }}
          locale={{ emptyText: '아직 입력된 VOC가 없습니다.' }}
          scroll={{ x: 900 }}
        />
      </Card>
    </Space>
  );
}
