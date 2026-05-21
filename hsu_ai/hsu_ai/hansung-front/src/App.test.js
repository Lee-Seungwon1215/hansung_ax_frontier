import { render, screen } from "@testing-library/react";
import { GoogleOAuthProvider } from "@react-oauth/google";
import App from "./App";

test("renders committee report builder", () => {
  render(
    <GoogleOAuthProvider clientId="test-client-id">
      <App />
    </GoogleOAuthProvider>
  );

  expect(screen.getByText("자료 기반 AI 보고서 변환")).toBeInTheDocument();
  expect(screen.getByText("보고서 초안 생성")).toBeInTheDocument();
});
