import { render, screen } from "@testing-library/react";
import { GoogleOAuthProvider } from "@react-oauth/google";
import App from "./App";

test("renders committee report builder", () => {
  render(
    <GoogleOAuthProvider clientId="test-client-id">
      <App />
    </GoogleOAuthProvider>
  );

  expect(screen.getByText("사업단운영위원회 주간업무보고")).toBeInTheDocument();
  expect(screen.getByText("운영위원회 보고서 생성")).toBeInTheDocument();
});
