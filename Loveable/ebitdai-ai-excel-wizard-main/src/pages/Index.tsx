
import { useState } from "react";
import { MessageCircle, CreditCard, User } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import ChatPanel from "@/components/ChatPanel";
import SubscriptionPanel from "@/components/SubscriptionPanel";

const Index = () => {
  const [activeTab, setActiveTab] = useState("chat");

  const tabs = [
    { id: "chat", icon: MessageCircle, label: "Chat" },
    { id: "subscription", icon: CreditCard, label: "Plans" },
  ];

  const renderContent = () => {
    switch (activeTab) {
      case "chat":
        return <ChatPanel />;
      case "subscription":
        return <SubscriptionPanel />;
      default:
        return <ChatPanel />;
    }
  };

  return (
    <div className="h-screen bg-gray-50 flex">
      {/* Sidebar */}
      <div className="w-16 bg-white border-r border-gray-200 flex flex-col items-center py-4 shadow-sm">
        {/* Logo */}
        <div className="mb-8">
          <div className="w-10 h-10 bg-black rounded-lg flex items-center justify-center">
            <span className="text-white font-bold text-sm">E</span>
          </div>
        </div>

        {/* Navigation Tabs */}
        <div className="flex flex-col space-y-2 mb-auto">
          {tabs.map((tab) => {
            const Icon = tab.icon;
            return (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`w-12 h-12 rounded-lg flex items-center justify-center transition-all duration-200 ${
                  activeTab === tab.id
                    ? "bg-black text-white shadow-lg"
                    : "text-gray-600 hover:bg-gray-100 hover:text-gray-900"
                }`}
                title={tab.label}
              >
                <Icon size={20} />
              </button>
            );
          })}
        </div>

        {/* Sign In Button */}
        <Button
          variant="outline"
          size="sm"
          className="w-12 h-12 rounded-lg p-0 border-gray-300 hover:bg-gray-100"
          title="Sign In"
        >
          <User size={20} />
        </Button>
      </div>

      {/* Main Content */}
      <div className="flex-1 flex flex-col">
        {/* Header */}
        <div className="bg-white border-b border-gray-200 px-6 py-4 shadow-sm">
          <div className="flex items-center">
            <div>
              <h1 className="text-2xl font-bold text-black">
                EBITD<span className="text-red-600">AI</span>
              </h1>
              <p className="text-sm text-gray-600 mt-1">
                Build Financial Models with AI
              </p>
            </div>
          </div>
        </div>

        {/* Content Area */}
        <div className="flex-1 p-6">
          {renderContent()}
        </div>
      </div>
    </div>
  );
};

export default Index;
