import { Alert, Button, Card, Col, Row, Space, Statistic, Table, Tag, Typography, Upload, message } from 'antd';
import type { ColumnsType } from 'antd/es/table';
import type { UploadFile } from 'antd/es/upload/interface';
import { ReloadOutlined, UploadOutlined } from '@ant-design/icons';
import axios from 'axios';
import { useEffect, useState } from 'react';

import { fetchBillingPreview, uploadBillingLoginFile, type BillingPreview, type BillingPreviewRow } from '../api/billing';

const statusColors: Record<string, string> = {
  일치: 'success',
  '고객명 상이': 'warning',
  '실적 없음': 'error',
  '사업자번호 없음': 'default',
};

export function DashboardPage() {
  const [preview, setPreview] = useState<BillingPreview | null>(null);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [fileList, setFileList] = useState<UploadFile[]>([]);
  const [errorMessage, setErrorMessage] = useState('');

  const loadPreview = async () => {
    setLoading(true);
    setErrorMessage('');
    try {
      const data = await fetchBillingPreview();
      setPreview(data);
    } catch (error: unknown) {
      const detail = axios.isAxiosError(error) ? error.response?.data?.detail : undefined;
      const nextMessage = typeof detail === 'string' ? detail : '청구자료를 불러오지 못했습니다.';
      setErrorMessage(nextMessage);
      message.error(nextMessage);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadPreview();
  }, []);

  const handleUploadPreview = async () => {
    const file = fileList[0]?.originFileObj;
    if (!file) {
      message.warning('은행로그인실적파일(은행) 엑셀을 선택해주세요.');
      return;
    }
    setUploading(true);
    setErrorMessage('');
    try {
      const data = await uploadBillingLoginFile(file);
      setPreview(data);
      message.success('업로드한 은행 로그인 실적파일로 청구자료를 생성했습니다.');
    } catch (error: unknown) {
      const detail = axios.isAxiosError(error) ? error.response?.data?.detail : undefined;
      const nextMessage = typeof detail === 'string' ? detail : '업로드 파일로 청구자료를 생성하지 못했습니다.';
      setErrorMessage(nextMessage);
      message.error(nextMessage);
    } finally {
      setUploading(false);
    }
  };

  const columns: ColumnsType<BillingPreviewRow> = [
    { title: '구분', dataIndex: 'source_type', key: 'source_type', width: 80 },
    { title: '순번', dataIndex: 'sequence', key: 'sequence', width: 80 },
    { title: '고객번호', dataIndex: 'customer_number', key: 'customer_number', width: 120 },
    { title: '사업자번호', dataIndex: 'business_number', key: 'business_number', width: 130 },
    { title: '청구원본 고객명', dataIndex: 'billing_company_name', key: 'billing_company_name', width: 180 },
    { title: '실적파일 고객명', dataIndex: 'bank_company_name', key: 'bank_company_name', width: 180 },
    { title: '담당자', dataIndex: 'manager_name', key: 'manager_name', width: 100 },
    { title: '기준일자', dataIndex: 'base_date', key: 'base_date', width: 110 },
    { title: '최초로그인', dataIndex: 'first_login', key: 'first_login', width: 110 },
    { title: '최종로그인', dataIndex: 'latest_login', key: 'latest_login', width: 110 },
    { title: '로그인횟수', dataIndex: 'login_count', key: 'login_count', align: 'right', width: 100 },
    {
      title: '상태',
      dataIndex: 'match_status',
      key: 'match_status',
      width: 120,
      render: (value) => <Tag color={statusColors[value] ?? 'default'}>{value}</Tag>,
    },
    { title: '비고', dataIndex: 'note', key: 'note', width: 130 },
  ];

  return (
    <Space className="dashboard-page" direction="vertical" size={18}>
      <div className="page-heading">
        <div>
          <Typography.Title level={3}>청구자료 생성</Typography.Title>
          <Typography.Text type="secondary">
            청구 원본과 은행 로그인 실적파일을 대사해 청구자료 생성 전 확인 목록을 만듭니다.
          </Typography.Text>
        </div>
        <Button icon={<ReloadOutlined />} onClick={loadPreview} loading={loading}>
          새로고침
        </Button>
      </div>

      {errorMessage ? (
        <Alert
          showIcon
          type="warning"
          message="Google Sheet 접근 확인 필요"
          description={errorMessage}
        />
      ) : null}

      <Card title="은행로그인실적파일(은행) 업로드">
        <Space wrap>
          <Upload
            accept=".xlsx,.xls,.csv"
            beforeUpload={() => false}
            fileList={fileList}
            maxCount={1}
            onChange={({ fileList: nextFileList }) => setFileList(nextFileList.slice(-1))}
          >
            <Button icon={<UploadOutlined />}>엑셀 파일 선택</Button>
          </Upload>
          <Button type="primary" onClick={handleUploadPreview} loading={uploading} disabled={!fileList.length}>
            업로드 파일로 청구자료 생성
          </Button>
          <Typography.Text type="secondary">
            고객번호, 고객명, 최근로그인, 로그인 컬럼을 기준으로 청구원본과 대사합니다.
          </Typography.Text>
        </Space>
      </Card>

      <Row gutter={[16, 16]}>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="전체" value={preview?.summary.total_count ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="개설" value={preview?.summary.open_count ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="연계" value={preview?.summary.erp_count ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="일치" value={preview?.summary.matched_count ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="고객명 상이" value={preview?.summary.name_mismatch_count ?? 0} suffix="건" />
          </Card>
        </Col>
        <Col xs={24} md={8} lg={4}>
          <Card>
            <Statistic title="확인 필요" value={preview?.summary.missing_count ?? 0} suffix="건" />
          </Card>
        </Col>
      </Row>

      <Card
        title={preview?.spreadsheet_title ?? '청구자료대사'}
        extra={
          preview?.spreadsheet_url ? (
            <a href={preview.spreadsheet_url} target="_blank" rel="noreferrer">
              원본 시트
            </a>
          ) : null
        }
      >
        <Table
          rowKey={(row) => `${row.source_type}-${row.sequence}-${row.customer_number}-${row.business_number}`}
          columns={columns}
          dataSource={preview?.rows ?? []}
          loading={loading}
          scroll={{ x: 1600 }}
          pagination={{ pageSize: 20 }}
        />
      </Card>
    </Space>
  );
}
