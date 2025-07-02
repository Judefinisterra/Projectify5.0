
import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { Textarea } from "@/components/ui/textarea";
import { Send, Mic, ArrowUp, Plus } from "lucide-react";

const ChatPanel = () => {
  const [message, setMessage] = useState("");

  const handleSend = () => {
    if (message.trim()) {
      console.log("Sending message:", message);
      setMessage("");
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className="h-full flex flex-col max-w-4xl mx-auto">
      {/* Welcome Message */}
      <div className="flex items-center justify-center mt-12">
        <Card className="w-full max-w-2xl border-0 shadow-none bg-transparent">
          <CardContent className="text-center py-12">
            <h2 className="text-4xl font-light text-gray-800 mb-6">
              What do you want to model?
            </h2>
            <div className="space-y-3 text-gray-600">
              <div className="flex items-center justify-center gap-3">
                <span className="bg-black text-white rounded-full w-6 h-6 flex items-center justify-center text-sm font-medium">1</span>
                <span>Tell us about your business idea or financial scenario</span>
              </div>
              <div className="flex items-center justify-center gap-3">
                <span className="bg-black text-white rounded-full w-6 h-6 flex items-center justify-center text-sm font-medium">2</span>
                <span>We'll create a professional Excel model for you</span>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Input Area */}
      <div className="bg-white rounded-full border border-gray-200 shadow-sm p-3 mt-4 max-w-2xl mx-auto w-full">
        <div className="flex items-center gap-3">
          <Button
            variant="ghost"
            size="sm"
            className="text-gray-500 hover:text-gray-700 hover:bg-transparent w-8 h-8 p-0 flex-shrink-0"
          >
            <Plus size={24} />
          </Button>
          
          <div className="flex-1 relative flex items-center">
            <Textarea
              placeholder="Describe your business model, revenue streams, or financial scenario..."
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              onKeyPress={handleKeyPress}
              className="min-h-[40px] resize-none border-0 focus:ring-0 focus:outline-none text-gray-800 placeholder-gray-500 py-2 flex items-center leading-none"
              style={{ boxShadow: "none" }}
            />
          </div>

          <div className="flex gap-2 flex-shrink-0">
            <Button
              variant="ghost"
              size="sm"
              className="text-gray-500 hover:text-gray-700 hover:bg-gray-100 w-10 h-10 p-0"
            >
              <Mic size={18} />
            </Button>
            
            <Button
              onClick={handleSend}
              disabled={!message.trim()}
              className="bg-black hover:bg-gray-800 text-white rounded-full w-10 h-10 p-0 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <ArrowUp size={18} />
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ChatPanel;
