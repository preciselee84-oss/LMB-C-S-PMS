import {
  Alert,
  Button,
  Card,
  Col,
  Form,
  Input,
  InputNumber,
  Modal,
  Row,
  Select,
  Space,
  Statistic,
  Switch,
  Table,
  Tabs,
  Tag,
  Typography,
  message,
} from 'antd';
import type { ColumnsType } from 'antd/es/table';
import {
  BankOutlined,
  CheckCircleOutlined,
  DeleteOutlined,
  EditOutlined,
  FileDoneOutlined,
  ReloadOutlined,
  SendOutlined,
  ThunderboltOutlined,
  UserAddOutlined,
} from '@ant-design/icons';
import { useEffect, useMemo, useState } from 'react';

import {
  approveAdvanceRequest,
  createAdvanceRequest,
  createBankAccount,
  createWorkplace,
  fetchWorkplaceDashboard,
  generateTransferRow,
  markRequestPaid,
  rejectAdvanceRequest,
  type AdvanceRequest,
  type BankAccount,
  type Workplace,
  type WorkplaceForecast,
  type WorkplaceSummary,
} from '../api/workplaces';
import {
  createUser,
  deleteUser,
  fetchUsers,
  updateUser,
  type AdminUser,
  type AdminUserCreate,
  type AdminUserUpdate,
} from '../api/users';

const won = new Intl.NumberFormat('ko-KR');

const statusColors: Record<string, string> = {
  요청: 'processing',
  '품의 확정': 'blue',
  반려: 'error',
  '이체 대상': 'warning',
  '이체 완료': 'success',
};

export function DashboardPage() {
  const [summary, setSummary] = useState<WorkplaceSummary | null>(null);
  const [workplaces, setWorkplaces] = useState<Workplace[]>([]);
  const [accounts, setAccounts] = useState<BankAccount[]>([]);
  const [requests, setRequests] = useState<AdvanceRequest[]>([]);
  const [users, setUsers] = useState<AdminUser[]>([]);
  const [editingUser, setEditingUser] = useState<AdminUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [workplaceForm] = Form.useForm();
  const [accountForm] = Form.useForm();
  const [requestForm] = Form.useForm();
  const [userForm] = Form.useForm();
  const [userEditForm] = Form.useForm();

  const loadDashboard = async () => {
    setLoading(true);
    try {
      const data = await fetchWorkplaceDashboard();
      const userRows = await fetchUsers();
      setSummary(data.summary);
      setWorkplaces(data.workplaces);
      setAccounts(data.accounts);
      setRequests(data.requests);
      setUsers(userRows);
    } catch {
      message.error('위탁사업장 데이터를 불러오지 못했습니다.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadDashboard();
  }, []);

  const workplaceOptions = workplaces.map((row) => ({
    label: row.workplace_name,
    value: row.id,
  }));

  const workplaceById = useMemo(
    () => new Map(workplaces.map((row) => [row.id, row.workplace_name])),
    [workplaces],
  );

  const workplaceColumns: ColumnsType<Workplace> = [
    { title: '사업장명', dataIndex: 'workplace_name', key: 'workplace_name' },
    { title: '사업자번호', dataIndex: 'business_number', key: 'business_number', responsive: ['md'] },
    { title: '정기 지급일', dataIndex: 'regular_payment_day', key: 'regular_payment_day', render: (v) => (v ? `${v}일` : '-') },
    { title: '담당자', dataIndex: 'manager_name', key: 'manager_name', responsive: ['md'] },
    { title: '연락처', dataIndex: 'manager_contact', key: 'manager_contact', responsive: ['lg'] },
  ];

  const accountColumns: ColumnsType<BankAccount> = [
    { title: '계좌명', dataIndex: 'account_name', key: 'account_name' },
    { title: '사업장', dataIndex: 'linked_workplace_id', key: 'linked_workplace_id', render: (v) => workplaceById.get(v) ?? '-' },
    { title: '은행', dataIndex: 'bank_name', key: 'bank_name', responsive: ['md'] },
    { title: '계좌번호', dataIndex: 'account_number', key: 'account_number' },
    { title: '잔액', dataIndex: 'balance', key: 'balance', align: 'right', render: (v) => `${won.format(v)}원` },
  ];

  const requestColumns: ColumnsType<AdvanceRequest> = [
    { title: '요청일시', dataIndex: 'requested_at', key: 'requested_at', render: (v) => new Date(v).toLocaleString('ko-KR') },
    { title: '사업장', dataIndex: 'workplace_name', key: 'workplace_name' },
    { title: '요청 금액', dataIndex: 'request_amount', key: 'request_amount', align: 'right', render: (v) => `${won.format(v)}원` },
    { title: '요청자', dataIndex: 'requested_by', key: 'requested_by', responsive: ['md'] },
    {
      title: '상태',
      dataIndex: 'status',
      key: 'status',
      render: (v) => <Tag color={statusColors[v] ?? 'default'}>{v}</Tag>,
    },
    {
      title: '처리',
      key: 'actions',
      render: (_, row) => (
        <Space wrap>
          <Button size="small" disabled={row.status !== '요청'} onClick={() => void handleApprove(row.id)}>
            승인
          </Button>
          <Button size="small" danger disabled={row.status !== '요청'} onClick={() => void handleReject(row.id)}>
            반려
          </Button>
          <Button size="small" disabled={row.status !== '품의 확정'} onClick={() => void handleTransfer(row.id)}>
            이체자료
          </Button>
          <Button size="small" disabled={!['품의 확정', '이체 대상'].includes(row.status)} onClick={() => void handlePaid(row.id)}>
            이체완료
          </Button>
        </Space>
      ),
    },
  ];

  const forecastColumns: ColumnsType<WorkplaceForecast> = [
    { title: '사업장', dataIndex: 'workplace_name', key: 'workplace_name' },
    { title: '월평균 지급액', dataIndex: 'average_monthly_amount', key: 'average_monthly_amount', align: 'right', render: (v) => `${won.format(v)}원` },
    { title: '추천 지급액', dataIndex: 'suggested_amount', key: 'suggested_amount', align: 'right', render: (v) => `${won.format(v)}원` },
    { title: '안내', dataIndex: 'guide', key: 'guide', responsive: ['md'] },
  ];

  const userColumns: ColumnsType<AdminUser> = [
    { title: 'ID', dataIndex: 'username', key: 'username' },
    { title: '성명', dataIndex: 'name', key: 'name' },
    { title: '이메일', dataIndex: 'email', key: 'email', responsive: ['md'] },
    {
      title: '권한',
      dataIndex: 'role',
      key: 'role',
      render: (value) => (value === 'admin' ? '관리자' : '사용자'),
    },
    {
      title: '상태',
      dataIndex: 'is_active',
      key: 'is_active',
      render: (value) => <Tag color={value ? 'success' : 'default'}>{value ? '사용' : '중지'}</Tag>,
    },
    {
      title: '관리',
      key: 'actions',
      render: (_, row) => (
        <Space wrap>
          <Button size="small" icon={<EditOutlined />} onClick={() => openUserEdit(row)}>
            수정
          </Button>
          <Button size="small" danger icon={<DeleteOutlined />} onClick={() => void handleDeleteUser(row)}>
            삭제
          </Button>
        </Space>
      ),
    },
  ];

  const handleCreateWorkplace = async (values: Partial<Workplace>) => {
    try {
      await createWorkplace(values);
      workplaceForm.resetFields();
      message.success('사업장 정보가 등록되었습니다.');
      await loadDashboard();
    } catch {
      message.error('사업장 등록에 실패했습니다. 중복 사업장명을 확인해주세요.');
    }
  };

  const handleCreateAccount = async (values: Partial<BankAccount>) => {
    try {
      await createBankAccount(values);
      accountForm.resetFields();
      message.success('계좌 정보가 등록되었습니다.');
      await loadDashboard();
    } catch {
      message.error('계좌 등록에 실패했습니다.');
    }
  };

  const handleCreateRequest = async (values: {
    workplace_id: number;
    request_amount: number;
    requested_by: string;
    request_reason?: string;
  }) => {
    try {
      await createAdvanceRequest(values);
      requestForm.resetFields();
      message.success('전도금 요청이 등록되었습니다.');
      await loadDashboard();
    } catch {
      message.error('전도금 요청 등록에 실패했습니다.');
    }
  };

  const handleApprove = async (id: number) => {
    await approveAdvanceRequest(id, '관리자');
    message.success('품의가 확정되었습니다.');
    await loadDashboard();
  };

  const handleReject = async (id: number) => {
    Modal.confirm({
      title: '전도금 요청 반려',
      content: '반려 처리하시겠습니까?',
      okText: '반려',
      okButtonProps: { danger: true },
      cancelText: '취소',
      onOk: async () => {
        await rejectAdvanceRequest(id, '관리자', '해커톤 시연 반려');
        message.success('요청이 반려되었습니다.');
        await loadDashboard();
      },
    });
  };

  const handleTransfer = async (id: number) => {
    try {
      const row = await generateTransferRow(id);
      Modal.info({
        title: 'CMS 이체자료 생성',
        content: (
          <div className="transfer-preview">
            <div>사업장: {row.workplace_name}</div>
            <div>은행: {row.bank_name}</div>
            <div>계좌번호: {row.account_number}</div>
            <div>예금주: {row.holder_name}</div>
            <div>이체금액: {won.format(row.amount)}원</div>
          </div>
        ),
      });
      await loadDashboard();
    } catch {
      message.error('연결된 사업장 계좌가 필요합니다.');
    }
  };

  const handlePaid = async (id: number) => {
    await markRequestPaid(id);
    message.success('이체 완료로 처리되었습니다.');
    await loadDashboard();
  };

  const handleCreateUser = async (values: AdminUserCreate) => {
    try {
      await createUser(values);
      userForm.resetFields();
      message.success('사용자가 등록되었습니다.');
      await loadDashboard();
    } catch {
      message.error('사용자 등록에 실패했습니다. 중복 ID를 확인해주세요.');
    }
  };

  const openUserEdit = (user: AdminUser) => {
    setEditingUser(user);
    userEditForm.setFieldsValue({
      name: user.name,
      email: user.email,
      role: user.role,
      is_active: user.is_active,
    });
  };

  const handleUpdateUser = async (values: AdminUserUpdate) => {
    if (!editingUser) {
      return;
    }
    const payload = { ...values };
    if (!payload.password) {
      delete payload.password;
    }
    await updateUser(editingUser.id, payload);
    setEditingUser(null);
    userEditForm.resetFields();
    message.success('사용자 정보가 수정되었습니다.');
    await loadDashboard();
  };

  const handleDeleteUser = async (user: AdminUser) => {
    Modal.confirm({
      title: '사용자 삭제',
      content: `${user.name}(${user.username}) 계정을 삭제하시겠습니까?`,
      okText: '삭제',
      okButtonProps: { danger: true },
      cancelText: '취소',
      onOk: async () => {
        await deleteUser(user.id);
        message.success('사용자가 삭제되었습니다.');
        await loadDashboard();
      },
    });
  };

  return (
    <Space className="dashboard-page" direction="vertical" size={18}>
      <div className="page-heading">
        <div>
          <Typography.Title level={3}>위탁사업장 관리 대시보드</Typography.Title>
          <Typography.Text type="secondary">사업장 등록부터 전도금 요청, 품의, 이체자료 생성까지 한 화면에서 처리합니다.</Typography.Text>
        </div>
        <Button icon={<ReloadOutlined />} onClick={loadDashboard} loading={loading}>
          새로고침
        </Button>
      </div>

      <Row gutter={[16, 16]}>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="등록 사업장" value={summary?.workplace_count ?? 0} suffix="곳" prefix={<BankOutlined />} />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="결재 대기" value={summary?.pending_count ?? 0} suffix="건" prefix={<FileDoneOutlined />} />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="이체 완료" value={summary?.paid_count ?? 0} suffix="건" prefix={<CheckCircleOutlined />} />
          </Card>
        </Col>
        <Col xs={24} md={6}>
          <Card>
            <Statistic title="누적 지급액" value={summary?.paid_amount ?? 0} formatter={(v) => `${won.format(Number(v))}원`} />
          </Card>
        </Col>
      </Row>

      <Tabs
        items={[
          {
            key: 'overview',
            label: '현황/예측',
            children: (
              <Row gutter={[16, 16]}>
                <Col xs={24} lg={16}>
                  <Card title="AI 지급 예측 MVP">
                    <Alert
                      showIcon
                      type="info"
                      icon={<ThunderboltOutlined />}
                      message="이체 완료 이력을 기준으로 월평균 지급액과 10% 여유분을 계산합니다."
                    />
                    <Table
                      className="section-table"
                      rowKey="workplace_id"
                      columns={forecastColumns}
                      dataSource={summary?.forecasts ?? []}
                      loading={loading}
                      pagination={false}
                    />
                  </Card>
                </Col>
                <Col xs={24} lg={8}>
                  <Card title="운영 지표">
                    <Space direction="vertical" className="full-width" size={14}>
                      <div className="metric-line">
                        <span>전체 요청</span>
                        <strong>{summary?.request_count ?? 0}건</strong>
                      </div>
                      <div className="metric-line">
                        <span>이번 달 요청</span>
                        <strong>{summary?.month_request_count ?? 0}건</strong>
                      </div>
                      <div className="metric-line">
                        <span>품의 확정</span>
                        <strong>{summary?.approved_count ?? 0}건</strong>
                      </div>
                      <div className="metric-line">
                        <span>계좌 잔액 합계</span>
                        <strong>{won.format(summary?.accounts_balance ?? 0)}원</strong>
                      </div>
                    </Space>
                  </Card>
                </Col>
              </Row>
            ),
          },
          {
            key: 'workplaces',
            label: '관리',
            children: (
              <Tabs
                items={[
                  {
                    key: 'workplace-info',
                    label: '사업장 정보관리',
                    children: (
                      <Row gutter={[16, 16]}>
                        <Col xs={24} lg={9}>
                          <Card title="사업장 등록">
                            <Form form={workplaceForm} layout="vertical" onFinish={handleCreateWorkplace}>
                              <Form.Item name="workplace_name" label="사업장명" rules={[{ required: true }]}>
                                <Input placeholder="예: 강남 위탁사업장" />
                              </Form.Item>
                              <Form.Item name="business_number" label="사업자번호">
                                <Input placeholder="000-00-00000" />
                              </Form.Item>
                              <Form.Item name="regular_payment_day" label="정기 지급일" initialValue={0}>
                                <InputNumber className="full-width" min={0} max={31} />
                              </Form.Item>
                              <Form.Item name="manager_name" label="담당자">
                                <Input placeholder="담당자명" />
                              </Form.Item>
                              <Form.Item name="manager_contact" label="담당자 연락처">
                                <Input placeholder="010-0000-0000" />
                              </Form.Item>
                              <Button type="primary" htmlType="submit" icon={<SendOutlined />} block>
                                사업장 등록
                              </Button>
                            </Form>
                          </Card>
                        </Col>
                        <Col xs={24} lg={15}>
                          <Card title="등록된 사업장">
                            <Table rowKey="id" columns={workplaceColumns} dataSource={workplaces} loading={loading} pagination={{ pageSize: 6 }} />
                          </Card>
                        </Col>
                      </Row>
                    ),
                  },
                  {
                    key: 'accounts',
                    label: '계좌 관리',
                    children: (
                      <Row gutter={[16, 16]}>
                        <Col xs={24} lg={9}>
                          <Card title="계좌 등록">
                            <Form form={accountForm} layout="vertical" onFinish={handleCreateAccount}>
                              <Form.Item name="linked_workplace_id" label="연결 사업장" rules={[{ required: true }]}>
                                <Select options={workplaceOptions} placeholder="사업장 선택" />
                              </Form.Item>
                              <Form.Item name="account_name" label="계좌명" rules={[{ required: true }]}>
                                <Input placeholder="예: 강남사업장 운영계좌" />
                              </Form.Item>
                              <Form.Item name="bank_name" label="은행" rules={[{ required: true }]}>
                                <Input placeholder="은행명" />
                              </Form.Item>
                              <Form.Item name="account_number" label="계좌번호" rules={[{ required: true }]}>
                                <Input placeholder="계좌번호" />
                              </Form.Item>
                              <Form.Item name="holder_name" label="예금주">
                                <Input placeholder="예금주" />
                              </Form.Item>
                              <Form.Item name="balance" label="현재 잔액" initialValue={0}>
                                <InputNumber className="full-width" min={0} step={100000} formatter={(v) => `${won.format(Number(v ?? 0))}`} />
                              </Form.Item>
                              <Button type="primary" htmlType="submit" block>
                                계좌 등록
                              </Button>
                            </Form>
                          </Card>
                        </Col>
                        <Col xs={24} lg={15}>
                          <Card title="계좌 현황">
                            <Table rowKey="id" columns={accountColumns} dataSource={accounts} loading={loading} pagination={{ pageSize: 6 }} />
                          </Card>
                        </Col>
                      </Row>
                    ),
                  },
                  {
                    key: 'users',
                    label: '사용자 관리',
                    children: (
                      <Row gutter={[16, 16]}>
                        <Col xs={24} lg={9}>
                          <Card title="사용자 등록">
                            <Form form={userForm} layout="vertical" onFinish={handleCreateUser} initialValues={{ role: 'staff', is_active: true }}>
                              <Form.Item name="username" label="ID" rules={[{ required: true }]}>
                                <Input placeholder="로그인 ID" />
                              </Form.Item>
                              <Form.Item name="password" label="초기 비밀번호" rules={[{ required: true, min: 4 }]}>
                                <Input.Password placeholder="초기 비밀번호" />
                              </Form.Item>
                              <Form.Item name="name" label="성명" rules={[{ required: true }]}>
                                <Input placeholder="사용자명" />
                              </Form.Item>
                              <Form.Item name="email" label="이메일">
                                <Input placeholder="user@example.com" />
                              </Form.Item>
                              <Form.Item name="role" label="권한" rules={[{ required: true }]}>
                                <Select
                                  options={[
                                    { label: '사용자', value: 'staff' },
                                    { label: '관리자', value: 'admin' },
                                  ]}
                                />
                              </Form.Item>
                              <Form.Item name="is_active" label="로그인 허용" valuePropName="checked">
                                <Switch checkedChildren="사용" unCheckedChildren="중지" />
                              </Form.Item>
                              <Button type="primary" htmlType="submit" icon={<UserAddOutlined />} block>
                                사용자 등록
                              </Button>
                            </Form>
                          </Card>
                        </Col>
                        <Col xs={24} lg={15}>
                          <Card title="사용자 목록">
                            <Table rowKey="id" columns={userColumns} dataSource={users} loading={loading} pagination={{ pageSize: 6 }} />
                          </Card>
                        </Col>
                      </Row>
                    ),
                  },
                ]}
              />
            ),
          },
          {
            key: 'requests',
            label: '전도금/결재',
            children: (
              <Row gutter={[16, 16]}>
                <Col xs={24} lg={8}>
                  <Card title="전도금 요청">
                    <Form form={requestForm} layout="vertical" onFinish={handleCreateRequest}>
                      <Form.Item name="workplace_id" label="사업장" rules={[{ required: true }]}>
                        <Select options={workplaceOptions} placeholder="사업장 선택" />
                      </Form.Item>
                      <Form.Item name="request_amount" label="요청 금액" rules={[{ required: true }]}>
                        <InputNumber className="full-width" min={1} step={100000} formatter={(v) => `${won.format(Number(v ?? 0))}`} />
                      </Form.Item>
                      <Form.Item name="requested_by" label="요청자" rules={[{ required: true }]} initialValue="현장 담당자">
                        <Input />
                      </Form.Item>
                      <Form.Item name="request_reason" label="요청 사유">
                        <Input.TextArea rows={3} placeholder="전도금 사용 목적 및 필요 사유" />
                      </Form.Item>
                      <Button type="primary" htmlType="submit" icon={<SendOutlined />} block>
                        요청 등록
                      </Button>
                    </Form>
                  </Card>
                </Col>
                <Col xs={24} lg={16}>
                  <Card title="전자결재 및 지급 이력">
                    <Table rowKey="id" columns={requestColumns} dataSource={requests} loading={loading} pagination={{ pageSize: 8 }} />
                  </Card>
                </Col>
              </Row>
            ),
          },
        ]}
      />
      <Modal
        title="사용자 정보 수정"
        open={Boolean(editingUser)}
        onCancel={() => setEditingUser(null)}
        onOk={() => userEditForm.submit()}
        okText="저장"
        cancelText="취소"
      >
        <Form form={userEditForm} layout="vertical" onFinish={handleUpdateUser}>
          <Form.Item name="name" label="성명" rules={[{ required: true }]}>
            <Input />
          </Form.Item>
          <Form.Item name="email" label="이메일">
            <Input />
          </Form.Item>
          <Form.Item name="password" label="새 비밀번호">
            <Input.Password placeholder="변경할 때만 입력" />
          </Form.Item>
          <Form.Item name="role" label="권한" rules={[{ required: true }]}>
            <Select
              options={[
                { label: '사용자', value: 'staff' },
                { label: '관리자', value: 'admin' },
              ]}
            />
          </Form.Item>
          <Form.Item name="is_active" label="로그인 허용" valuePropName="checked">
            <Switch checkedChildren="사용" unCheckedChildren="중지" />
          </Form.Item>
        </Form>
      </Modal>
    </Space>
  );
}
